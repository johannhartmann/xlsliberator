"""Tests for validation artifact models."""

from xlsliberator.ir_models import WorkbookIR
from xlsliberator.validation_models import (
    FormulaIR,
    GateExecutionStatus,
    SourceRef,
    UnsupportedArtifactIR,
    ValidationCertification,
    ValidationGateResult,
    ValidationSeverity,
    WorkbookArtifactIR,
)


def test_validation_models_serialize_to_json() -> None:
    """Validation models should serialize predictably."""
    workbook = WorkbookIR(file_path="book.xlsx", file_format="xlsx")
    source_ref = SourceRef(
        source_file="book.xlsx",
        sheet="Sheet1",
        cell_range="A1",
        artifact_type="formula",
        artifact_id="Sheet1!A1",
    )
    inventory = WorkbookArtifactIR(
        workbook=workbook,
        formulas=[FormulaIR(source_ref=source_ref, formula_text="=SUM(A1:A2)")],
        unsupported_artifacts=[
            UnsupportedArtifactIR(
                source_ref=source_ref,
                reason="unsupported test artifact",
                severity=ValidationSeverity.WARNING,
            )
        ],
    )

    json_text = inventory.model_dump_json()

    assert "book.xlsx" in json_text
    assert "unsupported test artifact" in json_text


def test_validation_model_defaults_are_safe() -> None:
    """Collection defaults should not be shared between instances."""
    workbook = WorkbookIR(file_path="book.xlsx", file_format="xlsx")
    first = WorkbookArtifactIR(workbook=workbook)
    second = WorkbookArtifactIR(workbook=workbook)

    first.metadata["changed"] = True

    assert second.metadata == {}
    assert second.formulas == []
    assert second.unsupported_artifacts == []


def test_all_required_passed_gates_can_certify() -> None:
    """Certification is derived from required canonical gate states."""
    gate = ValidationGateResult(
        gate_name="inventory",
        passed=True,
        message="inventory passed",
    )
    certification = ValidationCertification(gate_results=[gate])

    assert certification.certified
    assert certification.gate_results[0].severity == ValidationSeverity.INFO
    assert certification.warnings == []
    assert certification.errors == []


def test_gate_status_is_canonical_and_passed_is_projection() -> None:
    """Canonical status must override a contradictory legacy passed field."""
    gate = ValidationGateResult(
        gate_name="runtime",
        status=GateExecutionStatus.UNAVAILABLE,
        passed=True,
        message="runtime missing",
    )

    assert gate.status == GateExecutionStatus.UNAVAILABLE
    assert gate.passed is False


def test_certification_cannot_override_failed_required_gate() -> None:
    """A caller-provided certified flag cannot override canonical gates."""
    certification = ValidationCertification(
        certified=True,
        gate_results=[
            ValidationGateResult(
                gate_name="runtime",
                status=GateExecutionStatus.SKIPPED,
                message="not executed",
            )
        ],
    )

    assert certification.certified is False
