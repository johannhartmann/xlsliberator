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

    Uses test-and-fix loop with LLM retry logic to repair formulas.

    Args:
        excel_path: Path to original Excel file
        ods_path: Path to native-converted ODS file

    Returns:
        Statistics dictionary with keys:
        - formulas_scanned: Total formulas checked
        - formulas_needing_fix: Formulas with INDIRECT/ADDRESS issues
        - formulas_fixed: Successfully repaired formulas
        - formulas_failed: Failed repairs
        - repair_attempts: Total repair attempts made
        - avg_attempts_per_fix: Average attempts needed per successful fix

    Note:
        Excel's ADDRESS() accepts sheet name as 5th parameter:
        ADDRESS(row, col, abs, a1, "Sheet")

        LibreOffice Calc does NOT support this. Must convert to:
        OFFSET(Sheet.A1, row-1, col-1)
    """
    logger.info(f"Fixing INDIRECT/ADDRESS formulas: {excel_path.name} → {ods_path.name}")

    stats = {
        "formulas_scanned": 0,
        "formulas_needing_fix": 0,
        "formulas_fixed": 0,
        "formulas_failed": 0,
    }

    # Step 1: Build sheet name mapping (Excel → ODS with proper quoting)
    wb_excel = openpyxl.load_workbook(excel_path, data_only=False)

    # Extract sheet names from Excel
    excel_sheet_names = wb_excel.sheetnames

    # Get sheet names from ODS via UNO
    ods_sheet_names: list[str] = []
    with UnoCtx() as ctx:
        doc = ctx.desktop.loadComponentFromURL(f"file://{ods_path.absolute()}", "_blank", 0, ())
        sheets = doc.getSheets()
        for i in range(sheets.getCount()):
            sheet = sheets.getByIndex(i)
            ods_sheet_names.append(sheet.getName())
        doc.close(True)

    # Build mapping: Excel name → ODS quoted name
    # LibreOffice requires single quotes around sheet names with special chars
    sheet_name_mapping = {}
    for excel_name, ods_name in zip(excel_sheet_names, ods_sheet_names, strict=False):
        # Check if sheet name needs quoting (contains spaces, numbers, special chars)
        needs_quoting = (
            " " in ods_name
            or "-" in ods_name
            or ods_name[0].isdigit()
            or any(c in ods_name for c in "!@#$%^&*()+=[]{};:,.<>?/\\|`~")
        )

        quoted_name = f"'{ods_name}'" if needs_quoting else ods_name
        sheet_name_mapping[excel_name] = quoted_name

    logger.debug(f"Sheet name mapping: {sheet_name_mapping}")

    # Step 2: Scan ODS for broken formulas with INDIRECT/ADDRESS patterns
    # We scan the ODS (not Excel) because LibreOffice already converted syntax (commas→semicolons)

    formulas_to_fix: list[tuple[str, str, str]] = []  # (sheet, cell_addr, ods_formula)

    with UnoCtx() as ctx:
        doc = ctx.desktop.loadComponentFromURL(f"file://{ods_path.absolute()}", "_blank", 0, ())
        doc.calculateAll()  # Ensure formulas are calculated

        sheets = doc.getSheets()

        for sheet_idx in range(sheets.getCount()):
            sheet = sheets.getByIndex(sheet_idx)
            sheet_name = sheet.getName()

            # Get the corresponding Excel sheet to know which cells have formulas
            if sheet_name not in wb_excel.sheetnames:
                continue

            excel_sheet = wb_excel[sheet_name]

            for row in excel_sheet.iter_rows():
                for excel_cell in row:
                    stats["formulas_scanned"] += 1

                    # Skip non-formula cells
                    if not excel_cell.value or not isinstance(excel_cell.value, str):
                        continue
                    if not excel_cell.value.startswith("="):
                        continue

                    # Check if Excel formula has INDIRECT+ADDRESS (needs fixing)
                    excel_formula = excel_cell.value
                    if "INDIRECT" not in excel_formula or "ADDRESS" not in excel_formula:
                        continue

                    # Get the ODS cell
                    ods_cell = sheet.getCellByPosition(
                        excel_cell.column - 1,  # 0-indexed
                        excel_cell.row - 1,
                    )

                    # Check if it has an error (#NAME? = 525)
                    error = ods_cell.getError()
                    if error == 525:  # #NAME? error
                        ods_formula = ods_cell.getFormula()
                        formulas_to_fix.append((sheet_name, excel_cell.coordinate, ods_formula))
                        stats["formulas_needing_fix"] += 1

                        if stats["formulas_needing_fix"] <= 5:
                            logger.debug(
                                f"Found broken formula: {sheet_name}!{excel_cell.coordinate} (Error {error})"
                            )

        doc.close(True)

    wb_excel.close()

    if not formulas_to_fix:
        logger.info("No INDIRECT/ADDRESS formulas found needing repair")
        return stats

    logger.info(f"Found {len(formulas_to_fix)} formulas with INDIRECT/ADDRESS cross-sheet issues")

    # Step 3: Repair formulas using AST transformation (deterministic, no LLM)
    logger.info("Applying AST-based INDIRECT/ADDRESS → OFFSET transformation...")

    from xlsliberator.formula_ast_transformer import FormulaASTTransformer, FormulaTransformError

    transformer = FormulaASTTransformer(sheet_mapping=sheet_name_mapping)

    # Keep document open for applying fixes
    with UnoCtx() as ctx:
        doc = ctx.desktop.loadComponentFromURL(f"file://{ods_path.absolute()}", "_blank", 0, ())
        sheets = doc.getSheets()

        for sheet_name, cell_addr, ods_formula in formulas_to_fix:
            try:
                # Apply AST transformation
                fixed_formula = transformer.transform_indirect_address_to_offset(ods_formula)

                # Get UNO sheet and cell
                sheet = sheets.getByName(sheet_name)
                cell = sheet.getCellRangeByName(cell_addr)

                # Set the fixed formula
                cell.setFormula(fixed_formula)

                stats["formulas_fixed"] += 1

                if stats["formulas_fixed"] <= 5:
                    logger.debug(
                        f"Fixed {sheet_name}!{cell_addr}:\n"
                        f"  FROM: {ods_formula[:80]}...\n"
                        f"  TO:   {fixed_formula[:80]}..."
                    )

            except FormulaTransformError as e:
                stats["formulas_failed"] += 1
                logger.warning(f"Failed to transform {sheet_name}!{cell_addr}: {e}")
                # Leave original formula unchanged

        # Recalculate and save
        logger.info("Recalculating formulas...")
        doc.calculateAll()

        logger.info("Saving repaired formulas...")
        doc.store()
        doc.close(True)

    logger.success(
        f"Formula repair complete: {stats['formulas_fixed']}/{stats['formulas_needing_fix']} fixed, "
        f"{stats['formulas_failed']} failed"
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
