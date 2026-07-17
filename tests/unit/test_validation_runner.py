"""Tests for validation gate runner."""

import hashlib
import shutil
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

from click.testing import CliRunner

from xlsliberator.cli import cli
from xlsliberator.docker_runtime import DockerRuntimeIdentity, DockerRuntimeUnavailable
from xlsliberator.formula_certification import (
    FormulaCertificationResult,
    FormulaEvidenceRecord,
)
from xlsliberator.formula_semantics import build_formula_ir
from xlsliberator.ir_models import WorkbookIR
from xlsliberator.report import ConversionReport
from xlsliberator.scenarios.models import (
    Action,
    ActionKind,
    EnvironmentManifest,
    ObservationKind,
    ObservationRequest,
    ObservationValue,
    RuntimeIdentity,
    RuntimeTrace,
    Scenario,
    ScenarioStep,
    StepResult,
    ValueKind,
)
from xlsliberator.validation_models import (
    ArtifactCoverage,
    CanonicalArtifactIR,
    CertificationTier,
    GateExecutionStatus,
    SourceRef,
    TargetKind,
    TargetRef,
    UnsupportedArtifactIR,
    ValidationSeverity,
    WorkbookArtifactIR,
)
from xlsliberator.validation_runner import ValidationPlan, ValidationRunner


class FakeRuntime:
    def __init__(
        self,
        *,
        stages: dict[str, dict[str, Any]] | None = None,
        unavailable: bool = False,
        source_mutated: bool = False,
    ) -> None:
        self.stages = stages or {
            name: {"status": "passed", "error": None}
            for name in ("open", "recalculate", "save", "close", "reopen", "package")
        }
        self.unavailable = unavailable
        self.source_mutated = source_mutated

    def resolve_identity(self) -> DockerRuntimeIdentity:
        if self.unavailable:
            raise DockerRuntimeUnavailable("Docker missing; host fallback is disabled")
        return DockerRuntimeIdentity(
            image_reference="xlsliberator-libreoffice:26.2.4.2",
            image_id="sha256:fixed",
            version="26.2.4.2",
            architecture="arm64",
            probe={"uno_importable": True, "office_executable": "/opt/lo/soffice"},
        )

    def validate_document(self, _path: Path, *, image_id: str | None = None) -> dict[str, Any]:
        assert image_id == "sha256:fixed"
        return {
            "success": True,
            "data": {
                "stages": self.stages,
                "source_mutated": self.source_mutated,
                "container_name": "xlsliberator-lo-job",
                "runtime": {"profile_identifier": "unique-profile"},
            },
        }


def test_validation_runner_strict_fails_on_error_gate(tmp_path: Path, monkeypatch: Any) -> None:
    """Strict mode should fail certification on ERROR gates."""
    input_path = tmp_path / "book.xls"
    input_path.write_text("placeholder")
    inventory = WorkbookArtifactIR(
        workbook=WorkbookIR(file_path=str(input_path), file_format="xls"),
        unsupported_artifacts=[
            UnsupportedArtifactIR(
                source_ref=SourceRef(
                    source_file=str(input_path),
                    artifact_type="legacy_xls_biff",
                    artifact_id="legacy-xls-incomplete",
                ),
                reason="legacy XLS BIFF parsing incomplete",
                severity=ValidationSeverity.ERROR,
            )
        ],
    )

    import xlsliberator.validation_runner as runner_module

    monkeypatch.setattr(runner_module, "inspect_workbook", lambda _path: inventory)

    report = ValidationRunner(
        ValidationPlan(input_path=input_path, enabled_gates=["inventory"])
    ).run_all()

    assert not report.certification.certified
    assert report.certification.errors


def test_validation_runner_non_strict_does_not_weaken_certification(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """Non-strict mode changes exceptions, not certification semantics."""
    input_path = tmp_path / "book.xls"
    input_path.write_text("placeholder")
    inventory = WorkbookArtifactIR(
        workbook=WorkbookIR(file_path=str(input_path), file_format="xls"),
        unsupported_artifacts=[
            UnsupportedArtifactIR(
                source_ref=SourceRef(
                    source_file=str(input_path),
                    artifact_type="legacy_xls_biff",
                    artifact_id="legacy-xls-incomplete",
                ),
                reason="legacy XLS BIFF parsing incomplete",
                severity=ValidationSeverity.ERROR,
            )
        ],
    )

    import xlsliberator.validation_runner as runner_module

    monkeypatch.setattr(runner_module, "inspect_workbook", lambda _path: inventory)

    report = ValidationRunner(
        ValidationPlan(input_path=input_path, strict=False, enabled_gates=["inventory"])
    ).run_all()

    assert not report.certification.certified
    assert report.certification.gate_results[0].passed is False


def test_conversion_gate_rejects_failed_report(tmp_path: Path) -> None:
    """A failed conversion report must block certification."""
    output_path = tmp_path / "book.ods"
    output_path.write_bytes(b"not-an-ods")
    conversion = ConversionReport(
        input_file="book.xlsx", output_file=str(output_path), success=False
    )
    report = ValidationRunner(
        ValidationPlan(
            input_path=tmp_path / "book.xlsx",
            output_path=output_path,
            conversion_report=conversion,
            enabled_gates=["conversion"],
            strict=False,
        )
    ).run_all()

    assert not report.certification.certified
    assert report.certification.gate_results[0].status == GateExecutionStatus.FAILED


def test_conversion_gate_rejects_missing_output(tmp_path: Path) -> None:
    """A nominally successful report cannot certify without an ODS package."""
    output_path = tmp_path / "missing.ods"
    conversion = ConversionReport(
        input_file="book.xlsx", output_file=str(output_path), success=True
    )

    gate = ValidationRunner(
        ValidationPlan(
            input_path=tmp_path / "book.xlsx",
            output_path=output_path,
            conversion_report=conversion,
        )
    ).run_conversion_gate()

    assert gate.status == GateExecutionStatus.FAILED
    assert not gate.passed
    assert "missing" in gate.message


def test_successful_libreoffice_scenario_trace_is_a_certification_gate(tmp_path: Path) -> None:
    output = tmp_path / "book.ods"
    output.write_bytes(b"target workbook")
    digest = hashlib.sha256(output.read_bytes()).hexdigest()
    now = datetime.now(UTC)
    trace = RuntimeTrace(
        trace_id="target-trace",
        scenario_id="smoke",
        runtime_role="target",
        runtime_identity=RuntimeIdentity(
            runtime_kind="libreoffice_docker",
            runtime_version="26.2.4.2",
            image_digest="sha256:fixed",
        ),
        environment=EnvironmentManifest(),
        status=GateExecutionStatus.PASSED,
        started_at=now,
        ended_at=now,
        workbook_hash_before=digest,
        workbook_hash_after=digest,
    )

    report = ValidationRunner(
        ValidationPlan(
            input_path=tmp_path / "book.xlsx",
            output_path=output,
            target_trace=trace,
            enabled_gates=["target_scenario"],
            evidence_dir=tmp_path / "evidence",
        )
    ).run_all()

    assert report.certification.certified
    assert report.certification.tier is CertificationTier.LIBREOFFICE_RUNTIME_VALIDATED
    assert (tmp_path / "evidence" / "libreoffice-scenario-trace.json").is_file()


def test_target_scenario_gate_rejects_trace_for_different_output(tmp_path: Path) -> None:
    output = tmp_path / "book.ods"
    output.write_bytes(b"target workbook")
    now = datetime.now(UTC)
    trace = RuntimeTrace(
        trace_id="target-trace",
        scenario_id="smoke",
        runtime_role="target",
        runtime_identity=RuntimeIdentity(
            runtime_kind="libreoffice_docker",
            runtime_version="26.2.4.2",
            image_digest="sha256:fixed",
        ),
        environment=EnvironmentManifest(),
        status=GateExecutionStatus.PASSED,
        started_at=now,
        ended_at=now,
        workbook_hash_before="0" * 64,
        workbook_hash_after="0" * 64,
    )

    gate = ValidationRunner(
        ValidationPlan(input_path=tmp_path / "book.xlsx", output_path=output, target_trace=trace)
    ).run_target_scenario_gate()

    assert gate.status is GateExecutionStatus.FAILED
    assert "does not match" in gate.message


def test_artifact_coverage_gate_writes_source_target_and_diff_evidence(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    source_path = tmp_path / "book.xlsx"
    target_path = tmp_path / "book.ods"
    source_path.write_bytes(b"source")
    target_path.write_bytes(b"target")
    artifact = CanonicalArtifactIR(
        artifact_id="artifact:cell:one",
        family="cell",
        artifact_type="cell",
        locator="sheet/Sheet1/cell/A1",
        coverage=ArtifactCoverage.SEMANTIC,
        source_ref=SourceRef(
            source_file=str(source_path),
            artifact_type="cell",
            artifact_id="artifact:cell:one",
        ),
    )
    source = WorkbookArtifactIR(
        workbook=WorkbookIR(file_path=str(source_path), file_format="xlsx"),
        artifacts=[artifact],
    )
    target = WorkbookArtifactIR(
        inventory_role="target",
        workbook=WorkbookIR(file_path=str(target_path), file_format="ods"),
        artifacts=[artifact.model_copy(deep=True)],
    )
    import xlsliberator.validation_runner as runner_module

    monkeypatch.setattr(
        runner_module,
        "inspect_workbook",
        lambda _path, role="source": source if role == "source" else target,
    )
    evidence_dir = tmp_path / "evidence"
    gate = ValidationRunner(
        ValidationPlan(
            input_path=source_path,
            output_path=target_path,
            enabled_gates=["coverage"],
            evidence_dir=evidence_dir,
        )
    ).run_coverage_gate()

    assert gate.status is GateExecutionStatus.PASSED
    assert {path.name for path in evidence_dir.iterdir()} == {
        "source-inventory.json",
        "target-inventory.json",
        "inventory-diff.json",
    }


def test_skipped_required_gate_blocks_certification(tmp_path: Path) -> None:
    """A required skipped gate is never equivalent to passing."""
    report = ValidationRunner(
        ValidationPlan(
            input_path=tmp_path / "book.xlsx",
            enabled_gates=["macro"],
            strict=False,
        )
    ).run_all()

    assert not report.certification.certified
    assert report.certification.gate_results[0].status == GateExecutionStatus.SKIPPED


def test_parse_target_kind_rejects_removed_targets() -> None:
    """LibreOffice is the only accepted target."""
    from xlsliberator.validation_runner import parse_target_kind

    assert parse_target_kind("libreoffice") == [TargetKind.LIBREOFFICE]
    for unsupported in ("both", "openoffice"):
        try:
            parse_target_kind(unsupported)
        except ValueError:
            pass
        else:
            raise AssertionError(f"target should be rejected: {unsupported}")


def test_validate_cli_json(tmp_path: Path) -> None:
    """Validation CLI should emit JSON."""
    import openpyxl

    input_path = tmp_path / "book.xlsx"
    workbook = openpyxl.Workbook()
    active = workbook.active
    assert active is not None
    active["A1"] = "=1+1"
    workbook.save(input_path)
    workbook.close()

    result = CliRunner().invoke(
        cli,
        ["validate", str(input_path), "--json", "--non-strict"],
    )

    assert result.exit_code in {0, 1}
    assert '"gate_results"' in result.output


def test_validation_plan_repair_is_explicit(tmp_path: Path, monkeypatch: Any) -> None:
    """Repair metadata should only be enabled explicitly."""
    input_path = tmp_path / "book.xlsx"
    input_path.write_text("placeholder")
    inventory = WorkbookArtifactIR(
        workbook=WorkbookIR(file_path=str(input_path), file_format="xlsx"),
    )

    import xlsliberator.validation_runner as runner_module

    monkeypatch.setattr(runner_module, "inspect_workbook", lambda _path: inventory)

    report = ValidationRunner(
        ValidationPlan(input_path=input_path, repair=True, enabled_gates=["inventory"])
    ).run_all()

    assert report.certification.metadata["repair_enabled"] is True
    repair_gate = next(
        gate for gate in report.certification.gate_results if gate.gate_name == "repair"
    )
    assert repair_gate.passed
    assert repair_gate.details["attempt_count"] == 0


def _formula_repair_scenario() -> Scenario:
    return Scenario(
        id="formula-repair",
        steps=[
            ScenarioStep(
                id="recalculate",
                action=Action(kind=ActionKind.RECALCULATE),
                observations_after=[
                    ObservationRequest(
                        id="formula-result",
                        kind=ObservationKind.CELL,
                        selector={"sheet": "Sheet1", "address": "A1"},
                    ),
                    ObservationRequest(
                        id="stable-result",
                        kind=ObservationKind.CELL,
                        selector={"sheet": "Sheet1", "address": "B1"},
                    ),
                ],
            )
        ],
    )


def _formula_repair_trace(
    role: str,
    *,
    formula_value: int,
    stable_value: int,
    workbook_hash: str,
) -> RuntimeTrace:
    now = datetime.now(UTC)
    return RuntimeTrace(
        trace_id=f"{role}-{formula_value}-{stable_value}",
        scenario_id="formula-repair",
        runtime_role=role,  # type: ignore[arg-type]
        runtime_identity=RuntimeIdentity(
            runtime_kind="microsoft_excel" if role == "source" else "libreoffice_docker",
            runtime_version="test",
            image_digest="sha256:fixed" if role == "target" else None,
        ),
        environment=EnvironmentManifest(),
        status=GateExecutionStatus.PASSED,
        started_at=now,
        ended_at=now,
        workbook_hash_before=workbook_hash,
        workbook_hash_after=workbook_hash,
        steps=[
            StepResult(
                step_id="recalculate",
                action=ActionKind.RECALCULATE,
                status=GateExecutionStatus.PASSED,
                started_at=now,
                ended_at=now,
                observations_after={
                    "formula-result": ObservationValue(
                        kind=ValueKind.NUMBER,
                        value=formula_value,
                        formula='=INDIRECT(ADDRESS(1;2;4;1;"Sheet2"))',
                    ),
                    "stable-result": ObservationValue(
                        kind=ValueKind.NUMBER,
                        value=stable_value,
                    ),
                },
            )
        ],
    )


class _FormulaRepairRuntime:
    def request(self, payload: dict[str, Any]) -> dict[str, Any]:
        source = Path(str(payload["ods_path"]))
        destination = Path(str(payload["output_path"]))
        shutil.copy2(source, destination)
        destination.write_bytes(destination.read_bytes() + b"-repaired")
        return {"success": True, "data": {"applied": payload["formula_repairs"]}}


def _formula_repair_runner(
    tmp_path: Path,
    *,
    candidate_stable_value: int,
) -> tuple[ValidationRunner, bytes]:
    input_path = tmp_path / "book.xlsx"
    input_path.write_bytes(b"source")
    output = tmp_path / "book.ods"
    original = b"original-target"
    output.write_bytes(original)
    source_ref = SourceRef(
        source_file=str(input_path),
        sheet="Sheet1",
        cell_range="A1",
        artifact_type="formula",
        artifact_id="formula:Sheet1!A1",
    )
    target_ref = TargetRef(
        target_file=str(output),
        sheet="Sheet1",
        cell_range="A1",
        artifact_type="formula",
        artifact_id="formula:Sheet1!A1",
    )
    formula = '=INDIRECT(ADDRESS(1;2;4;1;"Sheet2"))'
    inventory = WorkbookArtifactIR(
        workbook=WorkbookIR(file_path=str(input_path), file_format="xlsx"),
        formulas=[build_formula_ir(source_ref=source_ref, formula=formula)],
    )
    original_hash = hashlib.sha256(original).hexdigest()
    source_trace = _formula_repair_trace(
        "source", formula_value=2, stable_value=7, workbook_hash="a" * 64
    )
    target_trace = _formula_repair_trace(
        "target", formula_value=1, stable_value=7, workbook_hash=original_hash
    )
    runner = ValidationRunner(
        ValidationPlan(
            input_path=input_path,
            output_path=output,
            repair=True,
            max_repair_iterations=2,
            scenario=_formula_repair_scenario(),
            source_trace=source_trace,
            target_trace=target_trace,
            evidence_dir=tmp_path / "evidence",
        ),
        runtime=_FormulaRepairRuntime(),  # type: ignore[arg-type]
    )
    runner._inventory = inventory
    runner._formula_result = FormulaCertificationResult(
        status=GateExecutionStatus.FAILED,
        formula_count=1,
        records=[
            FormulaEvidenceRecord(
                source_ref=source_ref,
                target_ref=target_ref,
                source_formula=formula,
                target_formula=formula,
                source_observation=source_trace.steps[0].observations_after["formula-result"],
                target_observation=target_trace.steps[0].observations_after["formula-result"],
                observation_id="formula-result",
                status=GateExecutionStatus.FAILED,
                errors=["runtime value/error differs"],
            )
        ],
        errors=["runtime value/error differs"],
    )

    class _CandidateScenarioRunner:
        def __init__(self, runtime: Any) -> None:
            assert runtime is runner.runtime

        def run(
            self,
            path: Path,
            environment: EnvironmentManifest,
            scenario: Scenario,
        ) -> RuntimeTrace:
            assert environment == source_trace.environment
            assert scenario.id == "formula-repair"
            candidate_hash = hashlib.sha256(path.read_bytes()).hexdigest()
            return _formula_repair_trace(
                "target",
                formula_value=2,
                stable_value=candidate_stable_value,
                workbook_hash=candidate_hash,
            )

    runner._candidate_scenario_runner = _CandidateScenarioRunner  # type: ignore[attr-defined]
    return runner, original


def test_formula_repair_commits_only_verified_candidate(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    runner, original = _formula_repair_runner(tmp_path, candidate_stable_value=7)
    import xlsliberator.validation_runner as runner_module

    candidate_runner = runner._candidate_scenario_runner  # type: ignore[attr-defined]
    monkeypatch.setattr(runner_module, "LibreOfficeScenarioRunner", candidate_runner)
    monkeypatch.setattr(
        runner_module,
        "certify_formulas",
        lambda *_args, **_kwargs: FormulaCertificationResult(
            status=GateExecutionStatus.PASSED,
            formula_count=1,
        ),
    )

    gate = runner.run_repair_gate()

    assert gate.status is GateExecutionStatus.PASSED
    assert gate.details["transaction"]["accepted_iteration"] == 1
    assert runner.plan.output_path is not None
    assert runner.plan.output_path.read_bytes() == original + b"-repaired"
    assert not list(tmp_path.glob(".book-formula-repair-*.ods"))


def test_formula_repair_rejects_regression_and_preserves_original(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    runner, original = _formula_repair_runner(tmp_path, candidate_stable_value=8)
    import xlsliberator.validation_runner as runner_module

    candidate_runner = runner._candidate_scenario_runner  # type: ignore[attr-defined]
    monkeypatch.setattr(runner_module, "LibreOfficeScenarioRunner", candidate_runner)
    monkeypatch.setattr(
        runner_module,
        "certify_formulas",
        lambda *_args, **_kwargs: FormulaCertificationResult(
            status=GateExecutionStatus.PASSED,
            formula_count=1,
        ),
    )

    gate = runner.run_repair_gate()

    assert gate.status is GateExecutionStatus.FAILED
    assert gate.details["transaction"]["regressions"] == ["recalculate:stable-result"]
    assert runner.plan.output_path is not None
    assert runner.plan.output_path.read_bytes() == original
    assert not list(tmp_path.glob(".book-formula-repair-*.ods"))


def test_macro_gate_fails_when_source_vba_not_embedded(tmp_path: Path, monkeypatch: Any) -> None:
    """A syntax-clean ODS with no embedded macros must fail when the source had VBA."""
    input_path = tmp_path / "book.xlsm"
    input_path.write_text("placeholder")
    output_path = tmp_path / "book.ods"
    output_path.write_text("placeholder-ods")

    inventory = WorkbookArtifactIR(
        workbook=WorkbookIR(file_path=str(input_path), file_format="xlsm"),
        metadata={"vba_modules": [{"name": "Module1", "procedures": ["Macro1", "Macro2"]}]},
    )

    class _Summary:
        total_modules = 0
        valid_syntax = 0
        syntax_errors = 0

    import xlsliberator.validation_runner as runner_module

    monkeypatch.setattr(runner_module, "inspect_workbook", lambda _path: inventory)
    monkeypatch.setattr(
        "xlsliberator.python_macro_manager.validate_all_embedded_macros",
        lambda _path: _Summary(),
    )

    gate = ValidationRunner(
        ValidationPlan(input_path=input_path, output_path=output_path, enabled_gates=["macro"])
    ).run_macro_gate()

    assert not gate.passed
    assert gate.severity == ValidationSeverity.ERROR
    assert gate.details["expected_vba_procedures"] == 2


def test_backend_gate_errors_when_no_backend(tmp_path: Path, monkeypatch: Any) -> None:
    """No discovered office backend must fail the backend gate with ERROR severity."""
    gate = ValidationRunner(
        ValidationPlan(input_path=tmp_path / "book.xlsx", enabled_gates=["backend"]),
        runtime=FakeRuntime(unavailable=True),  # type: ignore[arg-type]
    ).run_backend_gate()

    assert not gate.passed
    assert gate.status == GateExecutionStatus.UNAVAILABLE
    assert gate.severity == ValidationSeverity.ERROR


def test_target_runtime_stages_are_independent_and_evidenced(tmp_path: Path) -> None:
    output = tmp_path / "book.ods"
    output.write_bytes(b"placeholder")
    runtime = FakeRuntime(
        stages={
            "open": {"status": "passed", "error": None},
            "recalculate": {"status": "failed", "error": "calculation failed"},
            "save": {"status": "passed", "error": None},
            "close": {"status": "passed", "error": None},
            "reopen": {"status": "passed", "error": None},
            "package": {"status": "passed", "error": None},
        }
    )
    runner = ValidationRunner(
        ValidationPlan(input_path=tmp_path / "book.xlsx", output_path=output),
        runtime=runtime,  # type: ignore[arg-type]
    )

    recalc = runner.run_target_stage_gate("recalculate")
    save = runner.run_target_stage_gate("save")

    assert recalc.status == GateExecutionStatus.FAILED
    assert recalc.message == "calculation failed"
    assert save.status == GateExecutionStatus.PASSED
    evidence_path = Path(recalc.evidence_references[0])
    assert evidence_path.is_file()
    assert "sha256:fixed" in evidence_path.read_text()


def test_target_package_gate_rejects_source_mutation(tmp_path: Path) -> None:
    output = tmp_path / "book.ods"
    output.write_bytes(b"placeholder")
    runner = ValidationRunner(
        ValidationPlan(input_path=tmp_path / "book.xlsx", output_path=output),
        runtime=FakeRuntime(source_mutated=True),  # type: ignore[arg-type]
    )

    gate = runner.run_target_stage_gate("package")

    assert gate.status == GateExecutionStatus.FAILED
    assert "mutated" in gate.message
