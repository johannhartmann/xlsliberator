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


def fix_indirect_address_formulas(
    excel_path: Path,
    ods_path: Path,
) -> dict[str, int]:
    """Fix INDIRECT/ADDRESS cross-sheet formulas after native conversion.

    Args:
        excel_path: Path to original Excel file
        ods_path: Path to native-converted ODS file

    Returns:
        Statistics dictionary with keys:
        - formulas_scanned: Total formulas checked
        - formulas_needing_fix: Formulas with INDIRECT/ADDRESS issues
        - formulas_fixed: Successfully repaired formulas
        - formulas_failed: Failed repairs

    Note:
        Excel's ADDRESS() accepts sheet name as 5th parameter:
        ADDRESS(row, col, abs, a1, "Sheet")

        LibreOffice Calc does NOT support this. Must convert to:
        "$Sheet." & ADDRESS(row, col, abs, a1)
    """
    logger.info(f"Fixing INDIRECT/ADDRESS formulas: {excel_path.name} → {ods_path.name}")

    from xlsliberator.llm_formula_translator import LLMFormulaTranslator

    stats = {
        "formulas_scanned": 0,
        "formulas_needing_fix": 0,
        "formulas_fixed": 0,
        "formulas_failed": 0,
    }

    # Step 1: Scan Excel for formulas with INDIRECT/ADDRESS patterns
    wb_excel = openpyxl.load_workbook(excel_path, data_only=False)

    formulas_to_fix: list[tuple[str, str, str]] = []  # (sheet, cell_addr, formula)

    for sheet_name in wb_excel.sheetnames:
        sheet = wb_excel[sheet_name]

        for row in sheet.iter_rows():
            for cell in row:
                stats["formulas_scanned"] += 1

                if not cell.value or not isinstance(cell.value, str):
                    continue

                formula = cell.value

                # Check if formula contains INDIRECT and ADDRESS
                if "INDIRECT" in formula and "ADDRESS" in formula:
                    formulas_to_fix.append((sheet_name, cell.coordinate, formula))
                    stats["formulas_needing_fix"] += 1

                    if stats["formulas_needing_fix"] <= 5:
                        logger.debug(f"Found formula needing fix: {sheet_name}!{cell.coordinate}")

    wb_excel.close()

    if not formulas_to_fix:
        logger.info("No INDIRECT/ADDRESS formulas found needing repair")
        return stats

    logger.info(f"Found {len(formulas_to_fix)} formulas with INDIRECT/ADDRESS cross-sheet issues")

    # Step 2: Translate formulas using LLM
    translator = LLMFormulaTranslator()
    translated_formulas: list[tuple[str, str, str]] = []  # (sheet, cell, new_formula)

    for sheet_name, cell_addr, excel_formula in formulas_to_fix:
        try:
            calc_formula = translator.translate_excel_to_calc(
                excel_formula, issue_type="indirect_address_cross_sheet"
            )

            if calc_formula != excel_formula:
                translated_formulas.append((sheet_name, cell_addr, calc_formula))
                stats["formulas_fixed"] += 1

                if stats["formulas_fixed"] <= 3:
                    logger.debug(
                        f"Translated {sheet_name}!{cell_addr}:\n"
                        f"  FROM: {excel_formula[:80]}...\n"
                        f"  TO:   {calc_formula[:80]}..."
                    )
            else:
                stats["formulas_failed"] += 1
                logger.warning(f"Translation failed for {sheet_name}!{cell_addr}")

        except Exception as e:
            stats["formulas_failed"] += 1
            logger.warning(f"Error translating {sheet_name}!{cell_addr}: {e}")

    if not translated_formulas:
        logger.warning("No formulas were successfully translated")
        return stats

    logger.info(f"Successfully translated {len(translated_formulas)} formulas")

    # Step 3: Update formulas in ODS via UNO
    logger.info("Updating formulas in ODS...")

    with UnoCtx() as ctx:
        doc = ctx.desktop.loadComponentFromURL(f"file://{ods_path.absolute()}", "_blank", 0, ())

        sheets = doc.getSheets()
        updated_count = 0

        for sheet_name, cell_addr, new_formula in translated_formulas:
            try:
                sheet = sheets.getByName(sheet_name)
                cell_range = sheet.getCellRangeByName(cell_addr)
                cell_range.setFormula(new_formula)
                updated_count += 1

                if updated_count <= 3:
                    logger.debug(f"Updated {sheet_name}!{cell_addr}")

            except Exception as e:
                logger.warning(f"Failed to update {sheet_name}!{cell_addr}: {e}")

        # Recalculate all formulas
        logger.info("Recalculating formulas...")
        doc.calculateAll()

        # Save the document
        doc.store()
        doc.close(True)

    logger.success(
        f"Fixed {updated_count} formulas in {ods_path} (scanned: {stats['formulas_scanned']}, "
        f"needed fix: {stats['formulas_needing_fix']}, failed: {stats['formulas_failed']})"
    )

    return stats


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
    2. INDIRECT/ADDRESS cross-sheet formulas (Excel-specific syntax)
    """
    logger.info(f"Post-processing native ODS: {ods_path}")

    stats = {}

    # Fix 1: Named ranges
    stats["named_ranges_added"] = fix_named_ranges(excel_path, ods_path)

    # Fix 2: INDIRECT/ADDRESS formulas
    formula_stats = fix_indirect_address_formulas(excel_path, ods_path)
    stats.update(formula_stats)

    logger.success(f"Post-processing complete: {stats}")
    return stats
