"""Post-processor to fix LibreOffice native ODS conversion issues.

LibreOffice's native --convert-to ods conversion has known bugs:
1. Named ranges are NOT converted (0 named ranges in output)
2. This causes #NAME? errors in all formulas using named ranges

This module fixes these issues by:
1. Extracting named ranges from original Excel file
2. Adding them to the native-converted ODS file via UNO
"""

from pathlib import Path

import openpyxl
from loguru import logger

from xlsliberator.uno_conn import UnoCtx


def fix_named_ranges(
    excel_path: Path,
    ods_path: Path,
) -> int:
    """Fix named ranges in native-converted ODS file.

    Args:
        excel_path: Path to original Excel file
        ods_path: Path to native-converted ODS file

    Returns:
        Number of named ranges added

    Note:
        LibreOffice native conversion does NOT convert named ranges.
        This function extracts them from Excel and adds to ODS.
    """
    logger.info(f"Fixing named ranges: {excel_path.name} → {ods_path.name}")

    # Step 1: Extract named ranges from Excel
    wb_excel = openpyxl.load_workbook(excel_path, data_only=False)
    named_ranges_excel = {}

    for name, defn in wb_excel.defined_names.items():
        # Get the range definition
        destinations = list(defn.destinations)
        if destinations:
            sheet_name, cell_range = destinations[0]
            # Convert Excel reference to Calc reference format
            # Excel: Sheet1!$A$1:$B$10
            # Calc: $Sheet1.$A$1:$B$10
            calc_ref = f"${sheet_name}.{cell_range}"
            named_ranges_excel[name] = calc_ref
            logger.debug(f"Excel named range: {name} = {calc_ref}")

    wb_excel.close()

    if not named_ranges_excel:
        logger.info("No named ranges found in Excel")
        return 0

    logger.info(f"Found {len(named_ranges_excel)} named ranges in Excel")

    # Step 2: Add named ranges to ODS via UNO
    with UnoCtx() as ctx:
        # Open ODS file
        doc = ctx.desktop.loadComponentFromURL(f"file://{ods_path.absolute()}", "_blank", 0, ())

        # Get named ranges container
        named_ranges_ods = doc.getPropertyValue("NamedRanges")

        # Add each named range
        added = 0
        for name, calc_ref in named_ranges_excel.items():
            try:
                # Check if already exists
                if named_ranges_ods.hasByName(name):
                    logger.debug(f"Named range already exists: {name}")
                    continue

                # Add named range
                # Arguments: name, content, address (as Position), type (0=RANGE)
                from com.sun.star.table import CellAddress

                # Create a dummy cell address (0,0,0) since we're using content string
                addr = CellAddress()
                addr.Sheet = 0
                addr.Column = 0
                addr.Row = 0

                named_ranges_ods.addNewByName(name, calc_ref, addr, 0)
                added += 1
                logger.debug(f"Added named range: {name} = {calc_ref}")

            except Exception as e:
                logger.warning(f"Failed to add named range {name}: {e}")

        # Save the document
        doc.store()
        doc.close(True)

    logger.success(f"Added {added} named ranges to {ods_path}")
    return added


def post_process_native_ods(
    excel_path: Path,
    ods_path: Path,
) -> dict[str, int]:
    """Post-process native-converted ODS to fix known bugs.

    Args:
        excel_path: Path to original Excel file
        ods_path: Path to native-converted ODS file

    Returns:
        Dictionary with fix statistics

    Fixes applied:
    1. Named ranges (LibreOffice doesn't convert them)
    """
    logger.info(f"Post-processing native ODS: {ods_path}")

    stats = {}

    # Fix 1: Named ranges
    stats["named_ranges_added"] = fix_named_ranges(excel_path, ods_path)

    logger.success(f"Post-processing complete: {stats}")
    return stats
