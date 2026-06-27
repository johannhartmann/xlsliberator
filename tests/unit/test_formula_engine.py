"""Tests for formula validation seam."""

from xlsliberator.formula_engine import FormulaDialect, FormulaEngine
from xlsliberator.ir_models import WorkbookIR
from xlsliberator.validation_models import FormulaIR, SourceRef, WorkbookArtifactIR


def test_validate_formula_text_accepts_balanced_formula() -> None:
    """Balanced formulas should pass basic structural validation."""
    result = FormulaEngine().validate_formula_text("=SUM(1;2;3)", FormulaDialect.CALC_A1)

    assert result.success
    assert result.details["validation_scope"] == "basic_structural"


def test_validate_formula_text_rejects_missing_equals() -> None:
    """Formula text must start with equals."""
    result = FormulaEngine().validate_formula_text("SUM(1;2)", FormulaDialect.CALC_A1)

    assert not result.success
    assert result.error == "Formula must start with '='"


def test_validate_formula_text_rejects_unbalanced_parentheses() -> None:
    """Unbalanced parentheses should fail with structured errors."""
    result = FormulaEngine().validate_formula_text("=SUM(1;2", FormulaDialect.CALC_A1)

    assert not result.success
    assert result.error == "Unbalanced parentheses"


def test_validate_formula_text_rejects_unbalanced_quotes() -> None:
    """Unbalanced quotes should fail with structured errors."""
    result = FormulaEngine().validate_formula_text('="unterminated', FormulaDialect.CALC_A1)

    assert not result.success
    assert result.error == "Unbalanced quote"


def test_collect_formulas_from_inventory() -> None:
    """FormulaEngine should collect FormulaIR entries from artifact inventory."""
    workbook = WorkbookIR(file_path="book.xlsx", file_format="xlsx")
    formula = FormulaIR(
        source_ref=SourceRef(
            source_file="book.xlsx",
            sheet="Sheet1",
            cell_range="A1",
            artifact_type="formula",
            artifact_id="Sheet1!A1",
        ),
        formula_text="=1+1",
    )
    inventory = WorkbookArtifactIR(workbook=workbook, formulas=[formula])

    assert FormulaEngine().collect_formulas(inventory) == [formula]
