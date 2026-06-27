"""Gate-based validation runner."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

from xlsliberator.calc_backend import discover_backends
from xlsliberator.certification_report import CertificationReport
from xlsliberator.control_inventory import extract_controls_and_bindings_from_ods
from xlsliberator.formula_engine import FormulaDialect, FormulaEngine
from xlsliberator.formula_repair_loop import FormulaRepairLoop
from xlsliberator.inspect_workbook import inspect_workbook
from xlsliberator.validation_models import (
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
    target_kinds: list[TargetKind] = field(default_factory=lambda: [TargetKind.BOTH])
    strict: bool = True
    repair: bool = False
    max_repair_iterations: int = 0
    enabled_gates: list[str] = field(
        default_factory=lambda: ["inventory", "formula", "macro", "control", "backend"]
    )


class ValidationRunner:
    """Run validation gates and produce a certification report."""

    def __init__(self, plan: ValidationPlan) -> None:
        self.plan = plan
        self._inventory: WorkbookArtifactIR | None = None

    def run_inventory_gate(self) -> ValidationGateResult:
        """Run source workbook inventory gate."""
        try:
            self._inventory = inspect_workbook(self.plan.input_path)
        except Exception as exc:
            return ValidationGateResult(
                gate_name="inventory",
                passed=False,
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
            passed=passed,
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
        """Run basic source formula structural validation."""
        inventory = self._get_inventory()
        engine = FormulaEngine()
        failures = []
        for formula in engine.collect_formulas(inventory):
            try:
                dialect = FormulaDialect(formula.dialect)
            except ValueError:
                dialect = FormulaDialect.EXCEL_A1
            result = engine.validate_formula_text(formula.formula_text, dialect)
            if not result.success:
                failures.append(result)
        return ValidationGateResult(
            gate_name="formula",
            passed=not failures,
            severity=ValidationSeverity.ERROR if failures else ValidationSeverity.INFO,
            message=(
                "Formula structural validation passed"
                if not failures
                else f"{len(failures)} formula(s) failed structural validation"
            ),
            details={"failures": [failure.model_dump(mode="json") for failure in failures]},
        )

    def run_macro_gate(self) -> ValidationGateResult:
        """Run embedded macro syntax validation when an output ODS exists."""
        if self.plan.output_path is None or not self.plan.output_path.exists():
            return ValidationGateResult(
                gate_name="macro",
                passed=True,
                severity=ValidationSeverity.WARNING,
                message="Macro gate skipped because no output ODS exists",
            )

        try:
            from xlsliberator.python_macro_manager import validate_all_embedded_macros

            summary = validate_all_embedded_macros(self.plan.output_path)
        except Exception as exc:
            return ValidationGateResult(
                gate_name="macro",
                passed=False,
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
            passed=passed,
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
                passed=True,
                severity=ValidationSeverity.WARNING,
                message="Control gate skipped because no output ODS exists",
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
            passed=passed,
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
        """Run backend discovery gate.

        With no office backend, runtime validation cannot run, so the gate fails
        with ERROR severity. In strict mode this blocks certification — a green
        certificate must never be issued for a workbook that was never evaluated
        in a real office runtime.
        """
        backends = discover_backends()
        details = [backend.info.__dict__ for backend in backends]
        passed = bool(backends)
        return ValidationGateResult(
            gate_name="backend",
            passed=passed,
            severity=ValidationSeverity.INFO if passed else ValidationSeverity.ERROR,
            message=(
                f"Discovered {len(backends)} office backend(s)"
                if passed
                else "No LibreOffice or Apache OpenOffice backend discovered; "
                "runtime validation cannot be performed"
            ),
            details={"backends": details},
        )

    def run_repair_gate(self) -> ValidationGateResult:
        """Run deterministic formula repair planning when explicitly enabled."""
        if not self.plan.repair:
            return ValidationGateResult(
                gate_name="repair",
                passed=True,
                severity=ValidationSeverity.INFO,
                message="Repair gate disabled",
            )

        loop = FormulaRepairLoop()
        evidence = loop.collect_evidence_from_inventory(self._get_inventory())
        attempts = loop.propose_repairs(evidence)
        applied = (
            loop.apply_repairs_to_ods(self.plan.output_path, attempts)
            if self.plan.output_path is not None and self.plan.output_path.exists()
            else attempts
        )
        unresolved = [attempt for attempt in applied if not attempt.success]
        return ValidationGateResult(
            gate_name="repair",
            passed=not unresolved,
            severity=ValidationSeverity.WARNING if unresolved else ValidationSeverity.INFO,
            message=(
                "No deterministic repair evidence collected"
                if not applied
                else f"{len(applied)} deterministic repair attempt(s) planned"
            ),
            details={
                "evidence_count": len(evidence),
                "attempt_count": len(attempts),
                "applied_count": len(applied),
                "unresolved_count": len(unresolved),
                "attempts": [
                    {
                        "rule_name": attempt.rule_name,
                        "before": attempt.before,
                        "after": attempt.after,
                        "success": attempt.success,
                        "error": attempt.error,
                        "source_ref": attempt.evidence.source_ref.model_dump(mode="json"),
                        "target_ref": attempt.evidence.target_ref.model_dump(mode="json"),
                    }
                    for attempt in applied
                ],
            },
        )

    def run_all(self) -> CertificationReport:
        """Run all enabled gates and return certification report."""
        gate_map = {
            "inventory": self.run_inventory_gate,
            "formula": self.run_formula_gate,
            "macro": self.run_macro_gate,
            "control": self.run_control_gate,
            "backend": self.run_backend_gate,
            "repair": self.run_repair_gate,
        }
        enabled_gates = list(self.plan.enabled_gates)
        if self.plan.repair and "repair" not in enabled_gates:
            enabled_gates.append("repair")
        gate_results = [gate_map[name]() for name in enabled_gates if name in gate_map]
        unsupported = self._inventory.unsupported_artifacts if self._inventory else []
        strict_failures = [
            gate
            for gate in gate_results
            if not gate.passed
            and gate.severity in {ValidationSeverity.ERROR, ValidationSeverity.FATAL}
        ]
        certified = (
            not strict_failures
            if self.plan.strict
            else not any(
                gate.severity == ValidationSeverity.FATAL
                for gate in gate_results
                if not gate.passed
            )
        )
        certification = ValidationCertification(
            certified=certified,
            target_profiles=[target.value for target in self.plan.target_kinds],
            gate_results=gate_results,
            unsupported_artifacts=unsupported,
            warnings=[
                gate.message for gate in gate_results if gate.severity == ValidationSeverity.WARNING
            ],
            errors=[gate.message for gate in strict_failures],
            metadata={
                "input_path": str(self.plan.input_path),
                "output_path": str(self.plan.output_path) if self.plan.output_path else None,
                "strict": self.plan.strict,
                "repair_enabled": self.plan.repair,
                "max_repair_iterations": self.plan.max_repair_iterations,
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


def parse_target_kind(value: str) -> list[TargetKind]:
    """Parse CLI target value into target kinds."""
    normalized = value.lower()
    if normalized == "both":
        return [TargetKind.BOTH]
    if normalized == "libreoffice":
        return [TargetKind.LIBREOFFICE]
    if normalized == "openoffice":
        return [TargetKind.OPENOFFICE]
    raise ValueError(f"Unsupported target: {value}")
