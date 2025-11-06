#!/usr/bin/env python3
"""Test MATCH function behavior in LibreOffice Calc."""


from loguru import logger

from xlsliberator.uno_conn import UnoCtx, new_calc


def test_match_function() -> None:
    """Test MATCH function directly in Calc."""
    with UnoCtx() as uno_ctx:
        doc = new_calc(uno_ctx)
        sheet = doc.getSheets().getByIndex(0)  # type: ignore

        # Setup test data
        logger.info("Setting up test data...")

        # Put lookup value in A1
        cell_a1 = sheet.getCellByPosition(0, 0)  # type: ignore
        cell_a1.setValue(1)  # type: ignore

        # Put values in D1:D5
        values = [13, 16, 2, 1, 6]
        for i, val in enumerate(values):
            cell = sheet.getCellByPosition(3, i)  # type: ignore (column D = 3)
            cell.setValue(val)  # type: ignore

        # Test MATCH formula
        logger.info("Testing MATCH formula...")

        # Test 1: Direct MATCH
        cell_f1 = sheet.getCellByPosition(5, 0)  # type: ignore (column F)
        cell_f1.setFormula("=MATCH(A1,D1:D5,0)")  # type: ignore
        logger.info("Formula: =MATCH(A1,D1:D5,0)")
        logger.info(f"Result: {cell_f1.getValue()}")  # type: ignore

        # Test 2: MATCH with German function name
        cell_f2 = sheet.getCellByPosition(5, 1)  # type: ignore
        cell_f2.setFormula("=vergleich(A1;D1:D5;0)")  # type: ignore
        logger.info("\nFormula: =vergleich(A1;D1:D5;0)")
        logger.info(f"Result: {cell_f2.getValue()}")  # type: ignore

        # Test 3: IFERROR with MATCH
        cell_f3 = sheet.getCellByPosition(5, 2)  # type: ignore
        cell_f3.setFormula('=IFERROR(MATCH(A1,D1:D5,0),"")')  # type: ignore
        logger.info('\nFormula: =IFERROR(MATCH(A1,D1:D5,0),"")')
        logger.info(f"Result: {cell_f3.getValue()}")  # type: ignore
        logger.info(f"String: {cell_f3.getString()}")  # type: ignore

        # Test 4: IFERROR with German
        cell_f4 = sheet.getCellByPosition(5, 3)  # type: ignore
        cell_f4.setFormula('=IFERROR(vergleich(A1;D1:D5;0);"")')  # type: ignore
        logger.info('\nFormula: =IFERROR(vergleich(A1;D1:D5;0);"")')
        logger.info(f"Result: {cell_f4.getValue()}")  # type: ignore
        logger.info(f"String: {cell_f4.getString()}")  # type: ignore

        # Recalculate
        doc.calculateAll()  # type: ignore

        # Check results again after recalc
        logger.info("\n--- After recalculation ---")
        logger.info(f"F1 (MATCH): {cell_f1.getValue()}")  # type: ignore
        logger.info(f"F2 (vergleich): {cell_f2.getValue()}")  # type: ignore
        logger.info(f"F3 (IFERROR/MATCH): {cell_f3.getValue()} / '{cell_f3.getString()}'")  # type: ignore
        logger.info(f"F4 (IFERROR/vergleich): {cell_f4.getValue()} / '{cell_f4.getString()}'")  # type: ignore


if __name__ == "__main__":
    test_match_function()
