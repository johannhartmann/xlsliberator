#!/usr/bin/env python3
"""Test formula locale handling in LibreOffice."""

from loguru import logger

from xlsliberator.uno_conn import UnoCtx, new_calc


def test_formula_locale() -> None:
    """Test different ways of setting formulas with locales."""
    with UnoCtx() as uno_ctx:
        doc = new_calc(uno_ctx)
        sheet = doc.getSheets().getByIndex(0)  # type: ignore

        # Setup test data
        sheet.getCellByPosition(0, 0).setValue(1)  # type: ignore  # A1 = 1
        sheet.getCellByPosition(3, 3).setValue(1)  # type: ignore  # D4 = 1

        # Test 1: English formula with commas (international format)
        cell_f1 = sheet.getCellByPosition(5, 0)  # type: ignore
        cell_f1.setFormula("=MATCH(A1,D1:D5,0)")  # type: ignore
        logger.info("Test 1: English formula with commas")
        logger.info("  Set: =MATCH(A1,D1:D5,0)")
        logger.info(f"  Get: {cell_f1.getFormula()}")  # type: ignore
        logger.info(f"  Value: {cell_f1.getValue()}")  # type: ignore

        # Test 2: German formula with semicolons
        cell_f2 = sheet.getCellByPosition(5, 1)  # type: ignore
        cell_f2.setFormula("=vergleich(A1;D1:D5;0)")  # type: ignore
        logger.info("\nTest 2: German formula with semicolons")
        logger.info("  Set: =vergleich(A1;D1:D5;0)")
        logger.info(f"  Get: {cell_f2.getFormula()}")  # type: ignore
        logger.info(f"  Value: {cell_f2.getValue()}")  # type: ignore

        # Recalculate
        doc.calculateAll()  # type: ignore

        logger.info("\n--- After recalculation ---")
        logger.info(f"Test 1 value: {cell_f1.getValue()}")  # type: ignore
        logger.info(f"Test 2 value: {cell_f2.getValue()}")  # type: ignore


if __name__ == "__main__":
    test_formula_locale()
