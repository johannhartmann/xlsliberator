"""Tests for evidence-driven deterministic formula repair planning."""

from xlsliberator.formula_certification import (
    FormulaCertificationResult,
    FormulaEvidenceRecord,
)
from xlsliberator.formula_repair_loop import FormulaRepairEvidence, FormulaRepairLoop
from xlsliberator.scenarios.models import ObservationValue, ValueKind
from xlsliberator.validation_models import GateExecutionStatus, SourceRef, TargetRef


def _evidence(target_formula: str) -> FormulaRepairEvidence:
    return FormulaRepairEvidence(
        source_ref=SourceRef(
            source_file="book.xlsx",
            sheet="Sheet1",
            cell_range="A1",
            artifact_type="formula",
            artifact_id="Sheet1!A1",
        ),
        target_ref=TargetRef(
            target_file="book.ods",
            sheet="Sheet1",
            cell_range="A1",
            artifact_type="formula",
            artifact_id="Sheet1!A1",
        ),
        source_formula=target_formula,
        target_formula=target_formula,
        parse_error="parse failed",
    )


def test_repair_loop_creates_attempt_for_matching_rule() -> None:
    """Evidence with a matching rule should create a repair attempt."""
    attempts = FormulaRepairLoop().propose_repairs(
        [_evidence('=INDIRECT(ADDRESS(1;2;4;1;"Sheet2"))')]
    )

    assert len(attempts) == 1
    assert attempts[0].rule_name == "indirect_address"
    assert attempts[0].success


def test_repair_loop_reports_unresolved_without_matching_rule() -> None:
    """Evidence without a deterministic rule should be reportable."""
    attempts = FormulaRepairLoop().propose_repairs([_evidence("=SUM(A1:A2)")])

    assert len(attempts) == 1
    assert attempts[0].rule_name == "unresolved"
    assert not attempts[0].success


def test_repair_evidence_is_collected_only_from_failed_runtime_records() -> None:
    seed = _evidence('=INDIRECT(ADDRESS(1;2;4;1;"Sheet2"))')
    record = FormulaEvidenceRecord(
        source_ref=seed.source_ref,
        target_ref=seed.target_ref,
        source_formula=seed.source_formula,
        target_formula=seed.target_formula,
        target_parser_result={"success": False, "error": {"message": "parse failed"}},
        source_observation=ObservationValue(kind=ValueKind.NUMBER, value=2),
        target_observation=ObservationValue(
            kind=ValueKind.ERROR,
            value="#REF!",
            error_type="#REF!",
        ),
        observation_id="result",
        dependencies=["Sheet2!B1"],
        spill_context={"spill_range": "A1:A2"},
        status=GateExecutionStatus.FAILED,
        errors=["target FormulaParser rejected formula", "runtime value/error differs"],
    )
    result = FormulaCertificationResult(
        status=GateExecutionStatus.FAILED,
        formula_count=1,
        records=[record],
        errors=record.errors,
    )

    evidence = FormulaRepairLoop().collect_evidence_from_result(
        result,
        scenario_id="exact-scenario",
    )

    assert len(evidence) == 1
    assert evidence[0].target_parser_result == record.target_parser_result
    assert evidence[0].source_value == record.source_observation.model_dump(mode="json")
    assert evidence[0].target_value == record.target_observation.model_dump(mode="json")
    assert evidence[0].dependencies == ("Sheet2!B1",)
    assert evidence[0].spill_context == {"spill_range": "A1:A2"}
    assert evidence[0].scenario_id == "exact-scenario"
