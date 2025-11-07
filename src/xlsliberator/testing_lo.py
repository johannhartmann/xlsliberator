"""LibreOffice Calc formula equivalence testing utilities."""

import math
from pathlib import Path
from typing import Any

from loguru import logger
from openpyxl import load_workbook

from xlsliberator.uno_conn import UnoCtx, open_calc


class FormulaComparisonResult:
    """Results from comparing Excel and Calc formula values."""

    def __init__(self) -> None:
        """Initialize comparison result."""
        self.total_cells: int = 0
        self.formula_cells: int = 0
        self.matching: int = 0
        self.mismatching: int = 0
        self.errors: int = 0
        self.tolerance: float = 1e-9
        self.mismatches: list[dict[str, Any]] = []

    @property
    def match_rate(self) -> float:
        """Calculate percentage of matching formula values."""
        if self.formula_cells == 0:
            return 0.0
        return (self.matching / self.formula_cells) * 100

    def summary(self) -> str:
        """Generate summary report."""
        return f"""Formula Equivalence Test Results:
- Total cells: {self.total_cells:,}
- Formula cells: {self.formula_cells:,}
- Matching: {self.matching:,} ({self.match_rate:.2f}%)
- Mismatching: {self.mismatching:,}
- Errors: {self.errors:,}
- Tolerance: {self.tolerance}
"""


def values_equal(val1: Any, val2: Any, tolerance: float = 1e-9) -> bool:
    """Compare two cell values with tolerance for floating point numbers.

    Args:
        val1: First value (from Excel)
        val2: Second value (from Calc)
        tolerance: Absolute tolerance for numeric comparisons

    Returns:
        True if values are considered equal
    """

    # Handle empty values: Excel returns None for empty, Calc returns 0.0
    # Special case: formulas like =IFERROR(MATCH(...), "") return None in Excel, 0.0 in Calc
    if val1 is None and val2 == 0.0:
        return True
    if val1 == 0.0 and val2 is None:
        return True
    if val1 is None and val2 is None:
        return True

    # If one is None (and other is not 0.0), they're not equal
    if val1 is None or val2 is None:
        return False

    # Handle numeric values
    if isinstance(val1, (int, float)) and isinstance(val2, (int, float)):
        # Check for NaN
        if math.isnan(val1) and math.isnan(val2):
            return True
        if math.isnan(val1) or math.isnan(val2):
            return False
        # Check for infinity
        if math.isinf(val1) and math.isinf(val2):
            return val1 == val2
        if math.isinf(val1) or math.isinf(val2):
            return False
        # Compare with tolerance
        return abs(val1 - val2) <= tolerance

    # Handle string values
    if isinstance(val1, str) and isinstance(val2, str):
        return val1.strip() == val2.strip()

    # Handle boolean values
    if isinstance(val1, bool) and isinstance(val2, bool):
        return val1 == val2

    # Type mismatch
    return False


def compare_excel_calc(
    excel_path: Path,
    ods_path: Path,
    tolerance: float = 1e-9,
) -> FormulaComparisonResult:
    """Compare formula values between Excel and LibreOffice Calc files.

    Args:
        excel_path: Path to original Excel file
        ods_path: Path to converted ODS file
        tolerance: Absolute tolerance for numeric comparisons

    Returns:
        FormulaComparisonResult with comparison statistics
    """
    result = FormulaComparisonResult()
    result.tolerance = tolerance

    logger.info(f"Comparing {excel_path} with {ods_path}")

    # Load Excel workbook - need to load twice to get both formulas and values
    logger.debug(f"Loading Excel workbook: {excel_path}")
    wb_excel_formulas = load_workbook(excel_path, data_only=False)  # For formulas
    wb_excel_values = load_workbook(excel_path, data_only=True)  # For calculated values

    # Connect to LibreOffice and open ODS
    logger.debug(f"Opening ODS file in LibreOffice: {ods_path}")
    with UnoCtx() as uno_ctx:
        doc = open_calc(uno_ctx, str(ods_path))

        # Force recalculation to ensure all formulas have current values
        logger.debug("Recalculating formulas in opened document...")
        doc.calculateAll()

        sheets = doc.getSheets()

        # Compare each sheet
        for sheet_name in wb_excel_formulas.sheetnames:
            if not sheets.hasByName(sheet_name):
                logger.warning(f"Sheet '{sheet_name}' not found in ODS file")
                continue

            logger.debug(f"Comparing sheet: {sheet_name}")
            ws_formulas = wb_excel_formulas[sheet_name]
            ws_values = wb_excel_values[sheet_name]
            sheet_calc = sheets.getByName(sheet_name)

            # Iterate through all cells in Excel sheet
            for _row_idx, row in enumerate(ws_formulas.iter_rows(), start=1):
                for cell in row:
                    result.total_cells += 1

                    # Skip non-formula cells
                    if (
                        cell.value is None
                        or not hasattr(cell, "data_type")
                        or cell.data_type != "f"
                    ):
                        continue

                    result.formula_cells += 1

                    # Get Excel calculated value from data_only workbook
                    value_cell = ws_values.cell(row=cell.row, column=cell.column)
                    excel_val = value_cell.value

                    # Get Calc value (zero-indexed)
                    calc_cell = sheet_calc.getCellByPosition(cell.column - 1, cell.row - 1)

                    # Check cell type to determine if we need string or numeric value
                    cell_type = calc_cell.getType().value  # TEXT, VALUE, FORMULA, EMPTY
                    if cell_type == "TEXT":
                        calc_val = calc_cell.getString()
                    else:
                        # For FORMULA and VALUE types, try numeric first, fallback to string
                        calc_val = calc_cell.getValue()
                        # If getValue returns 0 and cell has text, use text instead
                        if calc_val == 0.0:
                            calc_str = calc_cell.getString()
                            if calc_str and calc_str.strip():
                                calc_val = calc_str

                    # Compare values
                    if values_equal(excel_val, calc_val, tolerance):
                        result.matching += 1
                    else:
                        result.mismatching += 1
                        result.mismatches.append(
                            {
                                "sheet": sheet_name,
                                "cell": cell.coordinate,
                                "excel_value": excel_val,
                                "calc_value": calc_val,
                                "formula": cell.value if hasattr(cell, "value") else None,
                            }
                        )

                        if len(result.mismatches) <= 10:  # Log first 10 mismatches
                            logger.warning(
                                f"Mismatch at {sheet_name}!{cell.coordinate}: "
                                f"Excel={excel_val}, Calc={calc_val}"
                            )

    logger.success(result.summary())
    return result
