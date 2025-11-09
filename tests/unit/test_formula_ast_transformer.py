"""Unit tests for formula AST transformer.

Tests the transformation of INDIRECT/ADDRESS formulas and fixing
LibreOffice native conversion bugs.
"""

import pytest

from xlsliberator.formula_ast_transformer import (
    FormulaASTTransformer,
    FormulaTransformError,
)


class TestCrossSheetReferenceFixing:
    """Test fixing $SheetName.$Cell references (LibreOffice conversion bug)."""

    def test_simple_cross_sheet_reference(self):
        """Fix $Tabelle.$D$5 → Tabelle.$D$5"""
        transformer = FormulaASTTransformer()
        formula = "=SUM($Tabelle.$D$5)"

        # Expected: Remove $ before sheet name
        expected = "=SUM(Tabelle.$D$5)"
        result = transformer.transform_indirect_address_to_offset(formula)

        assert result == expected

    def test_cross_sheet_in_address_function(self):
        """Fix cross-sheet reference inside ADDRESS function."""
        transformer = FormulaASTTransformer()
        formula = "=ADDRESS(ROW()-27;$Tabelle.$D$5+1;4;1)"

        # Expected: Remove $ before sheet name
        expected = "=ADDRESS(ROW()-27;Tabelle.$D$5+1;4;1)"
        result = transformer.transform_indirect_address_to_offset(formula)

        assert result == expected

    def test_complex_indirect_with_cross_sheet(self):
        """Fix complex INDIRECT formula with cross-sheet reference."""
        transformer = FormulaASTTransformer()
        formula = '=MIN(INDIRECT("Spieler!"&ADDRESS(ROW()-27;2;4;1)&":"&ADDRESS(ROW()-27;$Tabelle.$D$5+1;4;1);1))'

        # Expected: Remove $ before Tabelle
        expected = '=MIN(INDIRECT("Spieler!"&ADDRESS(ROW()-27;2;4;1)&":"&ADDRESS(ROW()-27;Tabelle.$D$5+1;4;1);1))'
        result = transformer.transform_indirect_address_to_offset(formula)

        assert result == expected

    def test_multiple_cross_sheet_references(self):
        """Fix formula with multiple cross-sheet references."""
        transformer = FormulaASTTransformer()
        formula = "=$Tabelle.$A$1+$Tabelle.$B$2"

        # Expected: Remove $ before all sheet names
        expected = "=Tabelle.$A$1+Tabelle.$B$2"
        result = transformer.transform_indirect_address_to_offset(formula)

        assert result == expected

    def test_dont_break_valid_references(self):
        """Don't modify valid sheet references without leading $."""
        transformer = FormulaASTTransformer()
        formula = "=SUM(Tabelle.$D$5:Tabelle.$D$10)"

        # Expected: No change
        result = transformer.transform_indirect_address_to_offset(formula)

        assert result == formula


class TestIndirectAddressTransformation:
    """Test INDIRECT(ADDRESS(..., sheet)) transformations (original feature)."""

    def test_indirect_address_with_sheet_parameter(self):
        """Transform INDIRECT(ADDRESS(row, col, abs, a1, sheet)) pattern."""
        transformer = FormulaASTTransformer(sheet_mapping={"Sheet1": "Sheet1"})
        formula = '=INDIRECT(ADDRESS(5;3;4;1;"Sheet1"))'

        # Expected: Move sheet to concatenation
        expected = '=INDIRECT("Sheet1."&ADDRESS(5;3;4;1))'
        result = transformer.transform_indirect_address_to_offset(formula)

        assert result == expected

    def test_indirect_address_with_quoted_sheet(self):
        """Transform with sheet name that needs quoting."""
        transformer = FormulaASTTransformer(sheet_mapping={"My Sheet": "'My Sheet'"})
        formula = '=INDIRECT(ADDRESS(1;1;4;1;"My Sheet"))'

        # Expected: Use quoted sheet name
        expected = "=INDIRECT(\"'My Sheet'.\"&ADDRESS(1;1;4;1))"
        result = transformer.transform_indirect_address_to_offset(formula)

        assert result == expected

    def test_indirect_without_address(self):
        """Don't transform INDIRECT without ADDRESS."""
        transformer = FormulaASTTransformer()
        formula = '=INDIRECT("A1")'

        # Expected: No change
        result = transformer.transform_indirect_address_to_offset(formula)

        assert result == formula

    def test_address_without_indirect(self):
        """Don't transform ADDRESS without INDIRECT."""
        transformer = FormulaASTTransformer()
        formula = '=ADDRESS(5;3;4;1;"Sheet1")'

        # Expected: No change
        result = transformer.transform_indirect_address_to_offset(formula)

        assert result == formula


class TestEdgeCases:
    """Test edge cases and error conditions."""

    def test_empty_formula(self):
        """Handle empty formula."""
        transformer = FormulaASTTransformer()

        with pytest.raises(FormulaTransformError):
            transformer.transform_indirect_address_to_offset("")

    def test_invalid_syntax(self):
        """Handle invalid formula syntax."""
        transformer = FormulaASTTransformer()

        with pytest.raises(FormulaTransformError):
            transformer.transform_indirect_address_to_offset("=SUM((((")

    def test_formula_without_equals(self):
        """Handle formula missing leading =."""
        transformer = FormulaASTTransformer()

        with pytest.raises(FormulaTransformError):
            transformer.transform_indirect_address_to_offset("SUM(A1:A10)")

    def test_nested_functions(self):
        """Handle deeply nested function calls."""
        transformer = FormulaASTTransformer()
        formula = "=IF(SUM(A1:A10)>0;AVERAGE(B1:B10);0)"

        # Expected: No change (no INDIRECT/ADDRESS)
        result = transformer.transform_indirect_address_to_offset(formula)

        assert result == formula


class TestRealWorldFormulas:
    """Test actual formulas from the Bundesliga spreadsheet."""

    def test_bundesliga_min_formula(self):
        """Test MIN formula from AM48."""
        transformer = FormulaASTTransformer()
        # Actual formula from LibreOffice native conversion
        formula = '=MIN(INDIRECT("Spieler!"&ADDRESS(ROW()-27;2;4;1)&":"&ADDRESS(ROW()-27;$Tabelle.$D$5+1;4;1);1))'

        # Expected: Fix $Tabelle → Tabelle
        expected = '=MIN(INDIRECT("Spieler!"&ADDRESS(ROW()-27;2;4;1)&":"&ADDRESS(ROW()-27;Tabelle.$D$5+1;4;1);1))'
        result = transformer.transform_indirect_address_to_offset(formula)

        assert result == expected

    def test_bundesliga_average_formula(self):
        """Test AVERAGE formula from AN48."""
        transformer = FormulaASTTransformer()
        formula = '=AVERAGE(INDIRECT("Spieler!"&ADDRESS(ROW()-27;2;4;1)&":"&ADDRESS(ROW()-27;$Tabelle.$D$5+1;4;1);1))'

        expected = '=AVERAGE(INDIRECT("Spieler!"&ADDRESS(ROW()-27;2;4;1)&":"&ADDRESS(ROW()-27;Tabelle.$D$5+1;4;1);1))'
        result = transformer.transform_indirect_address_to_offset(formula)

        assert result == expected

    def test_bundesliga_median_formula(self):
        """Test MEDIAN formula from AO48."""
        transformer = FormulaASTTransformer()
        formula = '=MEDIAN(INDIRECT("Spieler!"&ADDRESS(ROW()-27;2;4;1)&":"&ADDRESS(ROW()-27;$Tabelle.$D$5+1;4;1);1))'

        expected = '=MEDIAN(INDIRECT("Spieler!"&ADDRESS(ROW()-27;2;4;1)&":"&ADDRESS(ROW()-27;Tabelle.$D$5+1;4;1);1))'
        result = transformer.transform_indirect_address_to_offset(formula)

        assert result == expected


class TestPreserveFormulaStructure:
    """Test that transformer preserves formula structure."""

    def test_preserve_operators(self):
        """Preserve arithmetic operators."""
        transformer = FormulaASTTransformer()
        formula = "=A1+B1-C1*D1/E1^2"

        result = transformer.transform_indirect_address_to_offset(formula)

        assert result == formula

    def test_preserve_comparison_operators(self):
        """Preserve comparison operators."""
        transformer = FormulaASTTransformer()
        formula = "=IF(A1>B1;A1;B1)"

        result = transformer.transform_indirect_address_to_offset(formula)

        assert result == formula

    def test_preserve_concatenation(self):
        """Preserve string concatenation."""
        transformer = FormulaASTTransformer()
        formula = '="Hello"&" "&"World"'

        result = transformer.transform_indirect_address_to_offset(formula)

        assert result == formula

    def test_preserve_parentheses(self):
        """Preserve parentheses and precedence."""
        transformer = FormulaASTTransformer()
        formula = "=(A1+B1)*(C1+D1)"

        result = transformer.transform_indirect_address_to_offset(formula)

        assert result == formula
