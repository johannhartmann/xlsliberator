"""ODS file generation from IR using LibreOffice UNO."""

from typing import Any

from loguru import logger

from xlsliberator.ir_models import CellType, WorkbookIR
from xlsliberator.uno_conn import UnoCtx, new_calc, recalc


class ODSWriteError(Exception):
    """Raised when ODS writing fails."""


def build_calc_from_ir(
    ctx: UnoCtx,
    wb_ir: WorkbookIR,
    locale: str = "en-US",
) -> Any:
    """Build LibreOffice Calc document from WorkbookIR.

    Args:
        ctx: UNO connection context
        wb_ir: Workbook intermediate representation
        locale: Target locale for formulas ("en-US" or "de-DE")

    Returns:
        LibreOffice Calc document object

    Raises:
        ODSWriteError: If document creation fails

    Note:
        Phase F4 implementation - supports sheets, values, and 10 core formulas.
        Full feature support (named ranges, tables, charts) in later phases.
    """
    if not ctx.is_connected:
        raise ODSWriteError("Not connected to LibreOffice")

    logger.info(
        f"Building Calc document from IR: {wb_ir.sheet_count} sheets, "
        f"{wb_ir.total_cells} cells, {wb_ir.total_formulas} formulas"
    )

    try:
        # Create new Calc document
        doc = new_calc(ctx)

        # Get the first sheet (default Sheet1)
        sheets = doc.getSheets()
        first_sheet = sheets.getByIndex(0)

        # Process sheets - use single pass for now (two-pass causes timeout)
        for sheet_idx, sheet_ir in enumerate(wb_ir.sheets):
            if sheet_idx == 0:
                # Use the existing first sheet
                sheet = first_sheet
                sheet.setName(sheet_ir.name)
            else:
                # Create new sheet
                sheets.insertNewByName(sheet_ir.name, sheet_idx)
                sheet = sheets.getByName(sheet_ir.name)

            logger.debug(
                f"Processing sheet '{sheet_ir.name}': "
                f"{sheet_ir.cell_count} cells, {sheet_ir.formula_count} formulas"
            )

            # Write cells
            cells_written = 0
            formulas_written = 0

            for cell_ir in sheet_ir.cells:
                try:
                    # Get UNO cell
                    uno_cell = sheet.getCellByPosition(cell_ir.col, cell_ir.row)

                    # Write value based on cell type
                    if cell_ir.cell_type == CellType.FORMULA and cell_ir.formula:
                        # Use English/international format - LibreOffice handles locale
                        uno_cell.setFormula(cell_ir.formula)
                        formulas_written += 1
                    elif cell_ir.cell_type == CellType.NUMBER and cell_ir.value is not None:
                        uno_cell.setValue(float(cell_ir.value))
                    elif cell_ir.cell_type == CellType.STRING and cell_ir.value is not None:
                        uno_cell.setString(str(cell_ir.value))
                    elif cell_ir.cell_type == CellType.BOOLEAN and cell_ir.value is not None:
                        # UNO doesn't have a boolean type, use number (1/0)
                        uno_cell.setValue(1.0 if cell_ir.value else 0.0)
                    elif cell_ir.cell_type == CellType.ERROR:
                        # Write error as string for now
                        uno_cell.setString(str(cell_ir.value) if cell_ir.value else "#ERROR!")

                    cells_written += 1

                except Exception as e:
                    logger.warning(
                        f"Failed to write cell {cell_ir.address} on sheet '{sheet_ir.name}': {e}"
                    )
                    continue

            logger.debug(
                f"Sheet '{sheet_ir.name}': wrote {cells_written} cells, {formulas_written} formulas"
            )

        logger.success(
            f"Built Calc document: {wb_ir.sheet_count} sheets, formulas written successfully"
        )

        return doc

    except Exception as e:
        raise ODSWriteError(f"Failed to build Calc document: {e}") from e


def write_ods_from_ir(
    ctx: UnoCtx,
    wb_ir: WorkbookIR,
    output_path: str,
    locale: str = "en-US",
) -> None:
    """Write ODS file from WorkbookIR.

    Args:
        ctx: UNO connection context
        wb_ir: Workbook intermediate representation
        output_path: Output .ods file path
        locale: Target locale for formulas

    Raises:
        ODSWriteError: If write fails
    """
    from xlsliberator.uno_conn import save_as_ods

    doc = build_calc_from_ir(ctx, wb_ir, locale)

    try:
        # Recalculate before saving
        recalc(ctx, doc)
        logger.debug("Recalculated document before saving")

        # Save as ODS
        save_as_ods(ctx, doc, output_path)
        logger.success(f"Saved ODS file: {output_path}")

    finally:
        # Close document
        doc.close(True)
