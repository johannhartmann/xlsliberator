"""Gate-based validation runner."""

from __future__ import annotations

import hashlib
import json
import os
import tempfile
import zipfile
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any

from xlsliberator.artifact_inventory import (
    diff_inventories,
    disposition_coverage_errors,
)
from xlsliberator.certification_report import CertificationReport
from xlsliberator.control_inventory import extract_controls_and_bindings_from_ods
from xlsliberator.docker_runtime import DockerRuntimeUnavailable, LibreOfficeDockerRuntime
from xlsliberator.formula_certification import FormulaCertificationResult, certify_formulas
from xlsliberator.formula_repair_loop import FormulaRepairLoop
from xlsliberator.inspect_workbook import inspect_workbook
from xlsliberator.libreoffice_scenario_runner import LibreOfficeScenarioRunner
from xlsliberator.report import ConversionReport
from xlsliberator.scenarios.diff import diff_traces
from xlsliberator.scenarios.models import RuntimeTrace, Scenario
from xlsliberator.validation_models import (
    CertificationTier,
    GateExecutionStatus,
    InventoryDiff,
    TargetKind,
    ValidationCertification,
    ValidationGateResult,
    ValidationSeverity,
    WorkbookArtifactIR,
)


@dataclass
class ValidationPlan:
    """Validation plan configuration."""

    input_path: Path
    output_path: Path | None = None
    target_kinds: list[TargetKind] = field(default_factory=lambda: [TargetKind.LIBREOFFICE])
    strict: bool = True
    repair: bool = False
    max_repair_iterations: int = 0
    conversion_report: ConversionReport | None = None
    evidence_dir: Path | None = None
    scenario: Scenario | None = None
    source_trace: RuntimeTrace | None = None
    target_trace: RuntimeTrace | None = None
    enabled_gates: list[str] = field(
        default_factory=lambda: [
            "inventory",
            "coverage",
            "formula",
            "macro",
            "control",
            "runtime_identity",
            "backend",
            "target_open",
            "target_recalculate",
            "target_save",
            "target_close",
            "target_reopen",
            "target_package",
        ]
    )


class ValidationRunner:
    """Run validation gates and produce a certification report."""

    def __init__(
        self,
        plan: ValidationPlan,
        *,
        runtime: LibreOfficeDockerRuntime | None = None,
    ) -> None:
        self.plan = plan
        self.runtime = runtime or LibreOfficeDockerRuntime()
        self._inventory: WorkbookArtifactIR | None = None
        self._target_inventory: WorkbookArtifactIR | None = None
        self._inventory_diff: InventoryDiff | None = None
        self._formula_result: FormulaCertificationResult | None = None
        self._target_evidence: dict[str, Any] | None = None

    def run_inventory_gate(self) -> ValidationGateResult:
        """Run source workbook inventory gate."""
        try:
            self._inventory = inspect_workbook(self.plan.input_path)
        except Exception as exc:
            return ValidationGateResult(
                gate_name="inventory",
                status=GateExecutionStatus.FAILED,
                severity=ValidationSeverity.FATAL,
                message=f"Inventory failed: {exc}",
            )

        unsupported = self._inventory.unsupported_artifacts
        passed = not any(
            artifact.severity in {ValidationSeverity.ERROR, ValidationSeverity.FATAL}
            for artifact in unsupported
        )
        return ValidationGateResult(
            gate_name="inventory",
            status=GateExecutionStatus.PASSED if passed else GateExecutionStatus.FAILED,
            severity=ValidationSeverity.ERROR if not passed else ValidationSeverity.INFO,
            message=(
                "Inventory completed"
                if passed
                else f"Inventory found {len(unsupported)} unsupported artifact(s)"
            ),
            details={
                "formula_count": len(self._inventory.formulas),
                "unsupported_count": len(unsupported),
            },
        )

    def run_formula_gate(self) -> ValidationGateResult:
        """Require target parser round-trips and source/target runtime equivalence."""
        inventory = self._get_inventory()
        output = self.plan.output_path
        if output is None or not output.is_file():
            return ValidationGateResult(
                gate_name="formula",
                status=GateExecutionStatus.NOT_RUN,
                severity=ValidationSeverity.ERROR,
                message="Formula certification requires an existing target ODS",
            )
        self._formula_result = certify_formulas(
            inventory,
            output,
            self.plan.source_trace,
            self.plan.target_trace,
            self.plan.scenario,
            self.runtime,
        )
        evidence_path = self._evidence_dir() / "formula-evidence.json"
        evidence_path.parent.mkdir(parents=True, exist_ok=True)
        evidence_path.write_text(
            self._formula_result.model_dump_json(indent=2) + "\n",
            encoding="utf-8",
        )
        passed = self._formula_result.status is GateExecutionStatus.PASSED
        return ValidationGateResult(
            gate_name="formula",
            status=GateExecutionStatus.PASSED if passed else GateExecutionStatus.FAILED,
            severity=ValidationSeverity.INFO if passed else ValidationSeverity.ERROR,
            message=(
                f"All {self._formula_result.formula_count} formulas passed target parsing and runtime comparison"
                if passed
                else f"Formula certification failed with {len(self._formula_result.errors)} error(s)"
            ),
            details=self._formula_result.model_dump(mode="json"),
            evidence_references=[str(evidence_path)],
        )

    def run_coverage_gate(self) -> ValidationGateResult:
        """Require complete source-to-target artifact loss accounting."""
        source = self._get_inventory()
        output = self.plan.output_path
        if output is None or not output.is_file():
            return ValidationGateResult(
                gate_name="coverage",
                status=GateExecutionStatus.NOT_RUN,
                severity=ValidationSeverity.ERROR,
                message="Artifact coverage requires an existing target ODS",
            )
        try:
            self._target_inventory = inspect_workbook(output, role="target")
            self._inventory_diff = diff_inventories(source, self._target_inventory)
            source.dispositions = list(self._inventory_diff.dispositions)
            errors = disposition_coverage_errors(source)
            evidence = self._write_inventory_evidence()
        except Exception as exc:
            return ValidationGateResult(
                gate_name="coverage",
                status=GateExecutionStatus.FAILED,
                severity=ValidationSeverity.ERROR,
                message=f"Artifact coverage failed: {exc}",
            )
        return ValidationGateResult(
            gate_name="coverage",
            status=GateExecutionStatus.FAILED if errors else GateExecutionStatus.PASSED,
            severity=ValidationSeverity.ERROR if errors else ValidationSeverity.INFO,
            message=(
                f"Artifact coverage has {len(errors)} blocking disposition error(s)"
                if errors
                else f"All {len(source.artifacts)} source artifacts are accounted for"
            ),
            details={
                "source_artifact_count": len(source.artifacts),
                "target_artifact_count": len(self._target_inventory.artifacts),
                "matched_count": len(self._inventory_diff.matched),
                "missing_count": len(self._inventory_diff.missing_source_artifact_ids),
                "errors": errors,
            },
            evidence_references=evidence,
        )

    def run_macro_gate(self) -> ValidationGateResult:
        """Run embedded macro syntax validation when an output ODS exists."""
        if self.plan.output_path is None or not self.plan.output_path.exists():
            return ValidationGateResult(
                gate_name="macro",
                status=GateExecutionStatus.FAILED,
                severity=ValidationSeverity.ERROR,
                message="Macro validation failed because no output ODS exists",
            )

        try:
            from xlsliberator.python_macro_manager import validate_all_embedded_macros

            summary = validate_all_embedded_macros(self.plan.output_path)
        except Exception as exc:
            return ValidationGateResult(
                gate_name="macro",
                status=GateExecutionStatus.FAILED,
                severity=ValidationSeverity.ERROR,
                message=f"Macro validation failed: {exc}",
            )

        # Cross-check against the source workbook: a syntax-clean ODS with zero
        # embedded macros must not pass when the source actually contained VBA —
        # that means translation/embedding silently failed.
        expected_procedures = self._expected_vba_procedure_count()
        missing_embed = expected_procedures > 0 and summary.total_modules == 0

        passed = summary.syntax_errors == 0 and not missing_embed
        if missing_embed:
            message = (
                f"Source has {expected_procedures} VBA procedure(s) but no Python macros "
                "were embedded in the output ODS"
            )
        elif passed:
            message = "Embedded macro syntax validation passed"
        else:
            message = f"{summary.syntax_errors} embedded macro syntax error(s)"
        return ValidationGateResult(
            gate_name="macro",
            status=GateExecutionStatus.PASSED if passed else GateExecutionStatus.FAILED,
            severity=ValidationSeverity.ERROR if not passed else ValidationSeverity.INFO,
            message=message,
            details={
                "total_modules": summary.total_modules,
                "valid_syntax": summary.valid_syntax,
                "syntax_errors": summary.syntax_errors,
                "expected_vba_procedures": expected_procedures,
            },
        )

    def run_control_gate(self) -> ValidationGateResult:
        """Run ODS control/event inventory gate when an output ODS exists."""
        if self.plan.output_path is None or not self.plan.output_path.exists():
            return ValidationGateResult(
                gate_name="control",
                status=GateExecutionStatus.FAILED,
                severity=ValidationSeverity.ERROR,
                message="Control validation failed because no output ODS exists",
            )

        controls, event_bindings = extract_controls_and_bindings_from_ods(self.plan.output_path)
        controls_with_events = {
            binding.control_id for binding in event_bindings if binding.control_id
        }
        button_controls = [
            control for control in controls if "button" in control.control_type.lower()
        ]
        missing_handlers = [
            control.id for control in button_controls if control.id not in controls_with_events
        ]
        passed = not missing_handlers
        return ValidationGateResult(
            gate_name="control",
            status=GateExecutionStatus.PASSED if passed else GateExecutionStatus.FAILED,
            severity=ValidationSeverity.ERROR if missing_handlers else ValidationSeverity.INFO,
            message=(
                "Control inventory passed"
                if passed
                else f"{len(missing_handlers)} button control(s) lack event bindings"
            ),
            details={
                "control_count": len(controls),
                "event_binding_count": len(event_bindings),
                "missing_handlers": missing_handlers,
            },
        )

    def run_backend_gate(self) -> ValidationGateResult:
        """Require the configured, probed Docker backend."""
        evidence = self._get_target_evidence()
        identity = evidence.get("identity")
        passed = identity is not None
        return ValidationGateResult(
            gate_name="backend",
            status=(GateExecutionStatus.PASSED if passed else GateExecutionStatus.UNAVAILABLE),
            severity=ValidationSeverity.INFO if passed else ValidationSeverity.ERROR,
            message=(
                "Pinned LibreOffice Docker backend is available"
                if passed
                else str(evidence.get("identity_error") or "Docker backend unavailable")
            ),
            details={"identity": identity},
            evidence_references=self._target_evidence_references(),
        )

    def run_runtime_identity_gate(self) -> ValidationGateResult:
        """Require immutable image resolution and a matching PyUNO probe."""
        evidence = self._get_target_evidence()
        identity = evidence.get("identity")
        passed = identity is not None
        return ValidationGateResult(
            gate_name="runtime_identity",
            status=GateExecutionStatus.PASSED if passed else GateExecutionStatus.UNAVAILABLE,
            severity=ValidationSeverity.INFO if passed else ValidationSeverity.ERROR,
            message=(
                "Immutable image identity and PyUNO provenance probe passed"
                if passed
                else str(evidence.get("identity_error") or "Runtime identity unavailable")
            ),
            details={"identity": identity},
            evidence_references=self._target_evidence_references(),
        )

    def run_target_stage_gate(self, stage_name: str) -> ValidationGateResult:
        """Project one independently reported runtime stage into a required gate."""
        evidence = self._get_target_evidence()
        validation_error = evidence.get("validation_error")
        stage = dict((evidence.get("validation") or {}).get("stages", {}).get(stage_name) or {})
        raw_status = str(stage.get("status") or "not_run")
        try:
            status = GateExecutionStatus(raw_status)
        except ValueError:
            status = GateExecutionStatus.FAILED
            stage["error"] = f"Worker returned unknown status: {raw_status}"
        if validation_error and status == GateExecutionStatus.NOT_RUN:
            if evidence.get("output_missing"):
                status = GateExecutionStatus.FAILED
            elif evidence.get("identity") is None:
                status = GateExecutionStatus.UNAVAILABLE
            else:
                status = GateExecutionStatus.FAILED
        if stage_name == "package" and (evidence.get("validation") or {}).get("source_mutated"):
            status = GateExecutionStatus.FAILED
            stage["error"] = "Target validation mutated the source output"
        return ValidationGateResult(
            gate_name=f"target_{stage_name}",
            status=status,
            severity=(
                ValidationSeverity.INFO
                if status == GateExecutionStatus.PASSED
                else ValidationSeverity.ERROR
            ),
            message=(
                f"LibreOffice {stage_name} passed"
                if status == GateExecutionStatus.PASSED
                else str(stage.get("error") or validation_error or f"{stage_name} was not run")
            ),
            details=stage,
            evidence_references=self._target_evidence_references(),
        )

    def run_target_scenario_gate(self) -> ValidationGateResult:
        """Require a successful, immutable Docker LibreOffice scenario trace."""
        trace = self.plan.target_trace
        evidence_path = self._evidence_dir() / "libreoffice-scenario-trace.json"
        if trace is None:
            return ValidationGateResult(
                gate_name="target_scenario",
                status=GateExecutionStatus.NOT_RUN,
                severity=ValidationSeverity.ERROR,
                message="Required LibreOffice scenario trace was not supplied",
            )
        evidence_path.parent.mkdir(parents=True, exist_ok=True)
        evidence_path.write_text(trace.model_dump_json(indent=2) + "\n", encoding="utf-8")
        failures: list[str] = []
        if trace.runtime_role != "target":
            failures.append("trace runtime role is not target")
        if trace.runtime_identity.runtime_kind != "libreoffice_docker":
            failures.append("trace does not identify the Docker LibreOffice runtime")
        if not trace.runtime_identity.image_digest:
            failures.append("trace omits the immutable runtime image digest")
        if trace.status is not GateExecutionStatus.PASSED:
            failures.append(f"trace status is {trace.status.value}")
        if trace.workbook_hash_after != trace.workbook_hash_before:
            failures.append("trace reports workbook mutation")
        output = self.plan.output_path
        if output is None or not output.is_file():
            failures.append("certification output is missing")
        elif _hash_file(output) != trace.workbook_hash_before:
            failures.append("trace workbook hash does not match certification output")
        return ValidationGateResult(
            gate_name="target_scenario",
            status=GateExecutionStatus.FAILED if failures else GateExecutionStatus.PASSED,
            severity=ValidationSeverity.ERROR if failures else ValidationSeverity.INFO,
            message=(
                "; ".join(failures) if failures else "Docker LibreOffice scenario trace passed"
            ),
            details={
                "trace_id": trace.trace_id,
                "scenario_id": trace.scenario_id,
                "runtime_identity": trace.runtime_identity.model_dump(mode="json"),
            },
            evidence_references=[str(evidence_path)],
        )

    def run_repair_gate(self) -> ValidationGateResult:
        """Apply deterministic repairs transactionally and rerun the exact scenario."""
        if not self.plan.repair:
            return ValidationGateResult(
                gate_name="repair",
                status=GateExecutionStatus.SKIPPED,
                required=False,
                severity=ValidationSeverity.INFO,
                message="Repair gate disabled",
            )

        inventory = self._get_inventory()
        if not inventory.formulas:
            return ValidationGateResult(
                gate_name="repair",
                status=GateExecutionStatus.PASSED,
                severity=ValidationSeverity.INFO,
                message="No formula repair is required",
                details={"evidence_count": 0, "attempt_count": 0},
            )

        if self._formula_result is None:
            self.run_formula_gate()
        result = self._formula_result
        if result is None:
            return ValidationGateResult(
                gate_name="repair",
                status=GateExecutionStatus.FAILED,
                severity=ValidationSeverity.ERROR,
                message="Formula evidence is unavailable",
            )
        if result.status is GateExecutionStatus.PASSED:
            return ValidationGateResult(
                gate_name="repair",
                status=GateExecutionStatus.PASSED,
                severity=ValidationSeverity.INFO,
                message="No formula repair is required",
                details={"evidence_count": len(result.records), "attempt_count": 0},
            )
        loop = FormulaRepairLoop()
        evidence = loop.collect_evidence_from_result(
            result,
            scenario_id=self.plan.scenario.id if self.plan.scenario else None,
        )
        attempts = loop.propose_repairs(evidence)
        repairs = loop.worker_formula_repairs(attempts)
        unresolved = [attempt for attempt in attempts if not attempt.success]
        precondition_errors = self._repair_precondition_errors(repairs)
        if unresolved:
            precondition_errors.append(
                f"{len(unresolved)} formula failure(s) have no successful deterministic rule"
            )
        transaction: dict[str, Any] = {
            "schema_version": "1.0.0",
            "accepted": False,
            "precondition_errors": precondition_errors,
            "repair_count": len(repairs),
            "max_iterations": self.plan.max_repair_iterations,
            "iterations": [],
            "regressions": [],
        }
        if not precondition_errors:
            self._execute_formula_repair_transaction(loop, repairs, transaction)
        evidence_path = self._evidence_dir() / "formula-repair-transaction.json"
        evidence_path.parent.mkdir(parents=True, exist_ok=True)
        evidence_path.write_text(
            json.dumps(
                {
                    **transaction,
                    "evidence": [asdict(item) for item in evidence],
                    "attempts": [asdict(item) for item in attempts],
                },
                indent=2,
                sort_keys=True,
                default=_json_default,
            )
            + "\n",
            encoding="utf-8",
        )
        accepted = bool(transaction["accepted"])
        return ValidationGateResult(
            gate_name="repair",
            status=GateExecutionStatus.PASSED if accepted else GateExecutionStatus.FAILED,
            severity=ValidationSeverity.INFO if accepted else ValidationSeverity.ERROR,
            message=(
                f"Accepted {len(repairs)} deterministic repair(s) after exact-scenario verification"
                if accepted
                else "Formula repair transaction was not accepted"
            ),
            details={
                "evidence_count": len(evidence),
                "attempt_count": len(attempts),
                "applied_count": len(repairs) if accepted else 0,
                "unresolved_count": len(unresolved),
                "transaction": transaction,
            },
            evidence_references=[str(evidence_path)],
        )

    def _repair_precondition_errors(self, repairs: list[dict[str, str]]) -> list[str]:
        errors: list[str] = []
        if not repairs:
            errors.append("no located deterministic repairs are available")
        if self.plan.max_repair_iterations < 1:
            errors.append("max_repair_iterations must be at least 1")
        if self.plan.output_path is None or not self.plan.output_path.is_file():
            errors.append("target ODS is missing")
        if self.plan.scenario is None:
            errors.append("the exact failing scenario is missing")
        if self.plan.source_trace is None:
            errors.append("the source trace is missing")
        if self.plan.target_trace is None:
            errors.append("the pre-repair target trace is missing")
        return errors

    def _execute_formula_repair_transaction(
        self,
        loop: FormulaRepairLoop,
        repairs: list[dict[str, str]],
        transaction: dict[str, Any],
    ) -> None:
        output = self.plan.output_path
        scenario = self.plan.scenario
        source_trace = self.plan.source_trace
        target_trace = self.plan.target_trace
        assert output is not None and scenario is not None
        assert source_trace is not None and target_trace is not None
        original_hash = _hash_file(output)
        candidate_input = output
        pending_repairs = repairs
        candidates: list[Path] = []
        seen_plans: set[tuple[tuple[str, str, str], ...]] = set()
        try:
            for iteration_number in range(1, self.plan.max_repair_iterations + 1):
                plan_key = tuple(
                    sorted(
                        (item["sheet"], item["address"], item["formula"])
                        for item in pending_repairs
                    )
                )
                if not plan_key or plan_key in seen_plans:
                    transaction["stopped_reason"] = "repair plan is empty or repeated"
                    return
                seen_plans.add(plan_key)
                candidate = _temporary_ods_path(output)
                candidates.append(candidate)
                response = self.runtime.request(
                    {
                        "op": "apply_document_repairs",
                        "ods_path": str(candidate_input),
                        "output_path": str(candidate),
                        "formula_repairs": pending_repairs,
                        "named_ranges": [],
                    }
                )
                iteration: dict[str, Any] = {
                    "iteration": iteration_number,
                    "repairs": pending_repairs,
                    "candidate_worker_response": response,
                    "candidate_trace": None,
                    "candidate_formula_certification": None,
                    "regressions": [],
                }
                transaction["iterations"].append(iteration)
                if not response.get("success") or not candidate.is_file():
                    transaction["stopped_reason"] = "container repair operation failed"
                    return
                candidate_trace = LibreOfficeScenarioRunner(runtime=self.runtime).run(
                    candidate,
                    source_trace.environment,
                    scenario,
                )
                candidate_result = certify_formulas(
                    self._get_inventory(),
                    candidate,
                    source_trace,
                    candidate_trace,
                    scenario,
                    self.runtime,
                )
                before_diff = diff_traces(source_trace, target_trace, scenario)
                after_diff = diff_traces(source_trace, candidate_trace, scenario)
                before = {
                    (item.step_id, item.observation_id): item.matched
                    for item in before_diff.differences
                }
                regressions = [
                    f"{item.step_id}:{item.observation_id}"
                    for item in after_diff.differences
                    if before.get((item.step_id, item.observation_id), False) and not item.matched
                ]
                iteration["candidate_trace"] = candidate_trace.model_dump(mode="json")
                iteration["candidate_formula_certification"] = candidate_result.model_dump(
                    mode="json"
                )
                iteration["regressions"] = regressions
                transaction["regressions"] = regressions
                accepted = (
                    candidate_trace.status is GateExecutionStatus.PASSED
                    and candidate_result.status is GateExecutionStatus.PASSED
                    and after_diff.status is GateExecutionStatus.PASSED
                    and not regressions
                )
                if accepted:
                    if _hash_file(output) != original_hash:
                        transaction["precondition_errors"].append(
                            "target ODS changed concurrently during repair"
                        )
                        return
                    os.replace(candidate, output)
                    transaction["accepted"] = True
                    transaction["accepted_iteration"] = iteration_number
                    transaction["committed_sha256"] = _hash_file(output)
                    self.plan.target_trace = candidate_trace
                    self._formula_result = candidate_result
                    return
                next_evidence = loop.collect_evidence_from_result(
                    candidate_result,
                    scenario_id=scenario.id,
                )
                next_attempts = loop.propose_repairs(next_evidence)
                if any(not attempt.success for attempt in next_attempts):
                    transaction["stopped_reason"] = "a formula failure is not repairable"
                    return
                pending_repairs = loop.worker_formula_repairs(next_attempts)
                candidate_input = candidate
            transaction["stopped_reason"] = "maximum repair iterations exhausted"
        finally:
            for candidate in candidates:
                candidate.unlink(missing_ok=True)

    def run_conversion_gate(self) -> ValidationGateResult:
        """Require a successful report and a structurally valid ODS output."""
        report = self.plan.conversion_report
        output = self.plan.output_path
        details = {
            "report": None if report is None else json.loads(report.to_json()),
            "output_exists": bool(output and output.is_file()),
            "valid_zip": bool(output and output.is_file() and zipfile.is_zipfile(output)),
        }
        if report is None:
            return ValidationGateResult(
                gate_name="conversion",
                status=GateExecutionStatus.NOT_RUN,
                severity=ValidationSeverity.ERROR,
                message="Conversion did not return a report",
                details=details,
            )
        failures: list[str] = []
        if not report.success:
            failures.append("conversion report reports failure")
        if report.errors:
            failures.append("conversion report contains errors")
        if output is None or not output.is_file():
            failures.append("output file is missing")
        elif not zipfile.is_zipfile(output):
            failures.append("output is not a valid ODS ZIP package")
        return ValidationGateResult(
            gate_name="conversion",
            status=GateExecutionStatus.FAILED if failures else GateExecutionStatus.PASSED,
            severity=ValidationSeverity.ERROR if failures else ValidationSeverity.INFO,
            message="Conversion gate passed" if not failures else "; ".join(failures),
            details=details,
        )

    def run_all(self) -> CertificationReport:
        """Run all enabled gates and return certification report."""
        gate_map = {
            "conversion": self.run_conversion_gate,
            "inventory": self.run_inventory_gate,
            "coverage": self.run_coverage_gate,
            "formula": self.run_formula_gate,
            "macro": self.run_macro_gate,
            "control": self.run_control_gate,
            "backend": self.run_backend_gate,
            "runtime_identity": self.run_runtime_identity_gate,
            "target_open": lambda: self.run_target_stage_gate("open"),
            "target_recalculate": lambda: self.run_target_stage_gate("recalculate"),
            "target_save": lambda: self.run_target_stage_gate("save"),
            "target_close": lambda: self.run_target_stage_gate("close"),
            "target_reopen": lambda: self.run_target_stage_gate("reopen"),
            "target_package": lambda: self.run_target_stage_gate("package"),
            "target_scenario": self.run_target_scenario_gate,
            "repair": self.run_repair_gate,
        }
        enabled_gates = list(self.plan.enabled_gates)
        if self.plan.repair and "repair" not in enabled_gates:
            enabled_gates.append("repair")
        if self.plan.target_trace is not None and "target_scenario" not in enabled_gates:
            enabled_gates.append("target_scenario")
        gate_results = []
        for name in enabled_gates:
            gate = gate_map.get(name)
            if gate is None:
                gate_results.append(
                    ValidationGateResult(
                        gate_name=name,
                        status=GateExecutionStatus.NOT_RUN,
                        severity=ValidationSeverity.ERROR,
                        message=f"Required gate is not implemented: {name}",
                    )
                )
                continue
            try:
                gate_results.append(gate())
            except Exception as exc:
                gate_results.append(
                    ValidationGateResult(
                        gate_name=name,
                        status=GateExecutionStatus.FAILED,
                        severity=ValidationSeverity.ERROR,
                        message=f"Gate raised an exception: {exc}",
                        details={"exception_type": type(exc).__name__},
                    )
                )
        unsupported = self._inventory.unsupported_artifacts if self._inventory else []
        blocking_gates = [
            gate
            for gate in gate_results
            if gate.required and gate.status != GateExecutionStatus.PASSED
        ]
        certified = bool(gate_results) and not blocking_gates
        libreoffice_runtime_validated = any(
            gate.gate_name in {"target_scenario", "target_package"}
            and gate.status is GateExecutionStatus.PASSED
            for gate in gate_results
        )
        source_differential_validated = bool(
            self.plan.source_trace
            and self.plan.target_trace
            and self._formula_result
            and self._formula_result.formula_count > 0
            and self._formula_result.status is GateExecutionStatus.PASSED
            and any(
                gate.gate_name == "formula" and gate.status is GateExecutionStatus.PASSED
                for gate in gate_results
            )
        )
        certification = ValidationCertification(
            certified=certified,
            tier=(
                CertificationTier.SOURCE_DIFFERENTIAL_VALIDATED
                if certified and source_differential_validated
                else CertificationTier.LIBREOFFICE_RUNTIME_VALIDATED
                if certified and libreoffice_runtime_validated
                else CertificationTier.STRUCTURAL
            ),
            target_profiles=[target.value for target in self.plan.target_kinds],
            gate_results=gate_results,
            unsupported_artifacts=unsupported,
            warnings=[
                gate.message for gate in gate_results if gate.severity == ValidationSeverity.WARNING
            ],
            errors=[
                f"{gate.gate_name} [{gate.status.value}]: {gate.message}" for gate in blocking_gates
            ],
            metadata={
                "input_path": str(self.plan.input_path),
                "output_path": str(self.plan.output_path) if self.plan.output_path else None,
                "strict": self.plan.strict,
                "repair_enabled": self.plan.repair,
                "max_repair_iterations": self.plan.max_repair_iterations,
                "evidence_dir": str(self._evidence_dir()),
                "scenario_id": self.plan.scenario.id if self.plan.scenario else None,
                "source_trace_id": (
                    self.plan.source_trace.trace_id if self.plan.source_trace else None
                ),
                "target_trace_id": (
                    self.plan.target_trace.trace_id if self.plan.target_trace else None
                ),
                "conversion_report": (
                    json.loads(self.plan.conversion_report.to_json())
                    if self.plan.conversion_report is not None
                    else None
                ),
            },
        )
        return CertificationReport(certification=certification)

    def _get_inventory(self) -> WorkbookArtifactIR:
        if self._inventory is None:
            self.run_inventory_gate()
        if self._inventory is None:
            raise RuntimeError("Inventory is unavailable")
        return self._inventory

    def _expected_vba_procedure_count(self) -> int:
        """Count VBA procedures the source workbook contains, from the inventory."""
        try:
            inventory = self._get_inventory()
        except Exception:
            return 0
        modules = inventory.metadata.get("vba_modules", [])
        if not isinstance(modules, list):
            return 0
        return sum(
            len(module.get("procedures", []) or [])
            for module in modules
            if isinstance(module, dict)
        )

    def _evidence_dir(self) -> Path:
        if self.plan.evidence_dir is not None:
            return self.plan.evidence_dir
        anchor = self.plan.output_path or self.plan.input_path
        return anchor.parent / f"{anchor.stem}.evidence"

    def _get_target_evidence(self) -> dict[str, Any]:
        if self._target_evidence is not None:
            return self._target_evidence
        evidence: dict[str, Any] = {
            "target_kind": TargetKind.LIBREOFFICE.value,
            "identity": None,
            "identity_error": None,
            "validation": None,
            "validation_error": None,
            "output_missing": False,
        }
        output = self.plan.output_path
        if output is None or not output.is_file():
            evidence["output_missing"] = True
            evidence["validation_error"] = "Target validation output is missing"
        try:
            identity = self.runtime.resolve_identity()
            evidence["identity"] = asdict(identity)
        except DockerRuntimeUnavailable as exc:
            evidence["identity_error"] = str(exc)
            self._target_evidence = evidence
            self._write_target_evidence(evidence)
            return evidence

        if output is not None and output.is_file():
            try:
                response = self.runtime.validate_document(output, image_id=identity.image_id)
                if not response.get("success"):
                    error = response.get("error") or {}
                    evidence["validation_error"] = str(
                        error.get("message") or "Target worker failed"
                    )
                else:
                    evidence["validation"] = dict(response.get("data") or {})
            except DockerRuntimeUnavailable as exc:
                evidence["validation_error"] = str(exc)
        self._target_evidence = evidence
        self._write_target_evidence(evidence)
        return evidence

    def _write_target_evidence(self, evidence: dict[str, Any]) -> None:
        directory = self._evidence_dir()
        directory.mkdir(parents=True, exist_ok=True)
        (directory / "libreoffice-runtime.json").write_text(
            json.dumps(evidence, indent=2, sort_keys=True) + "\n"
        )

    def _write_inventory_evidence(self) -> list[str]:
        if (
            self._inventory is None
            or self._target_inventory is None
            or self._inventory_diff is None
        ):
            raise RuntimeError("source, target, and diff inventories are required")
        directory = self._evidence_dir()
        directory.mkdir(parents=True, exist_ok=True)
        paths = {
            "source-inventory.json": self._inventory.model_dump(mode="json"),
            "target-inventory.json": self._target_inventory.model_dump(mode="json"),
            "inventory-diff.json": self._inventory_diff.model_dump(mode="json"),
        }
        references = []
        for name, payload in paths.items():
            path = directory / name
            path.write_text(json.dumps(payload, indent=2, sort_keys=True) + "\n", encoding="utf-8")
            references.append(str(path))
        return references

    def _target_evidence_references(self) -> list[str]:
        return [str(self._evidence_dir() / "libreoffice-runtime.json")]


def parse_target_kind(value: str) -> list[TargetKind]:
    """Parse CLI target value into target kinds."""
    normalized = value.lower()
    if normalized == "libreoffice":
        return [TargetKind.LIBREOFFICE]
    raise ValueError(f"Unsupported target: {value}")


def _hash_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _temporary_ods_path(output: Path) -> Path:
    descriptor, candidate_name = tempfile.mkstemp(
        prefix=f".{output.stem}-formula-repair-",
        suffix=".ods",
        dir=output.parent,
    )
    os.close(descriptor)
    candidate = Path(candidate_name)
    candidate.unlink()
    return candidate


def _json_default(value: object) -> object:
    model_dump = getattr(value, "model_dump", None)
    if callable(model_dump):
        dumped: object = model_dump(mode="json")
        return dumped
    if isinstance(value, Path):
        return str(value)
    raise TypeError(f"Object of type {type(value).__name__} is not JSON serializable")
