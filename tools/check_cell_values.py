#!/usr/bin/env python3
"""Check actual cell values in problematic formulas."""

from pathlib import Path

from loguru import logger
from openpyxl import load_workbook

from xlsliberator.uno_conn import UnoCtx, open_calc


def check_cell_values(excel_path: Path, ods_path: Path) -> None:
    """Check cell values involved in formula calculations."""
    wb = load_workbook(excel_path, data_only=True)

    with UnoCtx() as uno_ctx:
        doc = open_calc(uno_ctx, str(ods_path))
        sheets = doc.getSheets()  # type: ignore

        sheet_name = "Spielplan"
        ws_excel = wb[sheet_name]
        sheet_calc = sheets.getByName(sheet_name)  # type: ignore

        # Check what's in A2
        logger.info("Cell A2 (lookup value):")
        logger.info(f"  Excel: {ws_excel['A2'].value}")
        a2_calc = sheet_calc.getCellByPosition(0, 1)  # type: ignore
        logger.info(
            f"  Calc:  {a2_calc.getString() if a2_calc.getType().value == 'TEXT' else a2_calc.getValue()}"
        )  # type: ignore

        # Check what's in the lookup range D2:D19
        logger.info("\nLookup range D2:D19:")
        logger.info("Excel values:")
        for row in range(2, 20):
            val = ws_excel[f"D{row}"].value
            if val:
                logger.info(f"  D{row}: {val}")

        logger.info("\nCalc values:")
        for row in range(1, 19):  # 0-indexed
            cell = sheet_calc.getCellByPosition(3, row)  # type: ignore (column D = 3)
            val = cell.getString() if cell.getType().value == "TEXT" else cell.getValue()  # type: ignore
            if val:
                logger.info(f"  D{row + 1}: {val}")


if __name__ == "__main__":
    test_data_dir = Path(__file__).parent.parent / "tests" / "data"
    excel_file = test_data_dir / "Bundesliga-Ergebnis-Tippspiel_V2.31_2025-26.xlsm"
    ods_file = Path("output.ods")

    check_cell_values(excel_file, ods_file)
