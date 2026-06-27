"""Tests for validation gate runner."""

from pathlib import Path
from typing import Any

from click.testing import CliRunner

from xlsliberator.cli import cli
from xlsliberator.ir_models import WorkbookIR
from xlsliberator.validation_models import (
    SourceRef,
    UnsupportedArtifactIR,
    ValidationSeverity,
    WorkbookArtifactIR,
)
from xlsliberator.validation_runner import ValidationPlan, ValidationRunner


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
    monkeypatch.setattr(runner_module, "discover_backends", lambda: [])

    report = ValidationRunner(
        ValidationPlan(input_path=input_path, enabled_gates=["inventory"])
    ).run_all()

    assert not report.certification.certified
    assert report.certification.errors


def test_validation_runner_non_strict_reports_without_error_cert_failure(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """Non-strict mode should allow ERROR gates short of FATAL."""
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

    assert report.certification.certified
    assert report.certification.gate_results[0].passed is False


def test_validate_cli_json(tmp_path: Path) -> None:
    """Validation CLI should emit JSON."""
    import openpyxl

    input_path = tmp_path / "book.xlsx"
    workbook = openpyxl.Workbook()
    workbook.active["A1"] = "=1+1"
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
    import xlsliberator.validation_runner as runner_module

    monkeypatch.setattr(runner_module, "discover_backends", lambda: [])

    gate = ValidationRunner(
        ValidationPlan(input_path=tmp_path / "book.xlsx", enabled_gates=["backend"])
    ).run_backend_gate()

    assert not gate.passed
    assert gate.severity == ValidationSeverity.ERROR
