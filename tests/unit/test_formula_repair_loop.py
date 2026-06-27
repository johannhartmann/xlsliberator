"""Tests for deterministic formula repair loop skeleton."""

from xlsliberator.formula_repair_loop import FormulaRepairEvidence, FormulaRepairLoop
from xlsliberator.validation_models import SourceRef, TargetRef


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
