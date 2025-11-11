"""MCP tool implementations for LibreOffice UNO operations.

Exposes xlsliberator functionality through Model Context Protocol (MCP) tools
for integration with Claude Agent SDK and other MCP clients.
"""

from pathlib import Path
from typing import Any

from loguru import logger

from xlsliberator.api import convert as convert_api
from xlsliberator.embed_macros import embed_python_macros
from xlsliberator.python_macro_manager import (
    enumerate_python_scripts,
    test_script_execution,
    validate_all_embedded_macros,
)
from xlsliberator.testing_lo import compare_excel_calc
from xlsliberator.uno_conn import UnoCtx, get_cell, get_sheet, open_calc, recalc

# ==============================================================================
# Document Operations
# ==============================================================================


async def convert_excel_to_ods(
    excel_path: str,
    output_path: str,
    embed_macros: bool = True,
    use_agent: bool = True,
) -> dict[str, Any]:
    """Convert Excel file to LibreOffice Calc ODS format.

    Args:
        excel_path: Path to input Excel file (.xlsx, .xlsm, .xlsb, .xls)
        output_path: Path for output ODS file
        embed_macros: If True, translate and embed VBA macros (default: True)
        use_agent: If True, use agent-based rewriting for complex VBA (default: True)

    Returns:
        Dictionary with conversion results:
        - success: bool
        - output_path: str
        - report: dict with conversion statistics
        - error: str (if failed)
    """
    try:
        logger.info(f"MCP: Converting {excel_path} to {output_path}")
        report = convert_api(
            input_path=Path(excel_path),
            output_path=Path(output_path),
            embed_macros=embed_macros,
            use_agent=use_agent,
        )

        return {
            "success": True,
            "output_path": str(output_path),
            "report": {
                "sheet_count": report.sheet_count,
                "total_cells": report.total_cells,
                "total_formulas": report.total_formulas,
                "vba_modules": report.vba_modules,
                "python_handlers": report.python_handlers,
                "duration_seconds": report.duration_seconds,
                "errors": len(report.errors),
                "warnings": len(report.warnings),
            },
        }
    except Exception as e:
        logger.error(f"MCP: Conversion failed: {e}")
        return {
            "success": False,
            "error": str(e),
        }


async def recalculate_document(ods_path: str) -> dict[str, Any]:
    """Force recalculation of all formulas in an ODS document.

    Args:
        ods_path: Path to ODS file

    Returns:
        Dictionary with:
        - success: bool
        - message: str
        - error: str (if failed)
    """
    try:
        logger.info(f"MCP: Recalculating {ods_path}")
        with UnoCtx() as ctx:
            doc = open_calc(ctx, ods_path)
            recalc(ctx, doc)
            doc.close(True)

        return {
            "success": True,
            "message": f"Recalculated {ods_path}",
        }
    except Exception as e:
        logger.error(f"MCP: Recalculation failed: {e}")
        return {
            "success": False,
            "error": str(e),
        }


# ==============================================================================
# Cell and Sheet Operations
# ==============================================================================


async def read_cell(ods_path: str, sheet_name: str, cell_address: str) -> dict[str, Any]:
    """Read cell value, formula, and type from an ODS document.

    Args:
        ods_path: Path to ODS file
        sheet_name: Sheet name (or index as string)
        cell_address: Cell address (e.g., 'A1', 'B10')

    Returns:
        Dictionary with:
        - success: bool
        - value: cell value (number, string, or None)
        - formula: cell formula (if formula cell)
        - type: cell type ('VALUE', 'TEXT', 'FORMULA', 'EMPTY')
        - error: str (if failed)
    """
    try:
        logger.debug(f"MCP: Reading {sheet_name}!{cell_address} from {ods_path}")
        with UnoCtx() as ctx:
            doc = open_calc(ctx, ods_path)
            sheet = get_sheet(ctx, doc, sheet_name)
            cell = get_cell(ctx, sheet, cell_address)

            cell_type = cell.getType().value
            value = None
            formula = None

            if cell_type == "TEXT":
                value = cell.getString()
            elif cell_type in ("VALUE", "FORMULA"):
                value = cell.getValue()
                # Try to get string if value is 0
                if value == 0.0:
                    cell_str = cell.getString()
                    if cell_str and cell_str.strip():
                        value = cell_str

            if cell_type == "FORMULA":
                formula = cell.getFormula()

            doc.close(True)

        return {
            "success": True,
            "value": value,
            "formula": formula,
            "type": cell_type,
        }
    except Exception as e:
        logger.error(f"MCP: Read cell failed: {e}")
        return {
            "success": False,
            "error": str(e),
        }


async def list_sheets(ods_path: str) -> dict[str, Any]:
    """List all sheet names in an ODS document.

    Args:
        ods_path: Path to ODS file

    Returns:
        Dictionary with:
        - success: bool
        - sheets: list of sheet names
        - count: number of sheets
        - error: str (if failed)
    """
    try:
        logger.debug(f"MCP: Listing sheets in {ods_path}")
        with UnoCtx() as ctx:
            doc = open_calc(ctx, ods_path)
            sheets = doc.getSheets()
            sheet_names = [sheets.getByIndex(i).getName() for i in range(sheets.getCount())]
            doc.close(True)

        return {
            "success": True,
            "sheets": sheet_names,
            "count": len(sheet_names),
        }
    except Exception as e:
        logger.error(f"MCP: List sheets failed: {e}")
        return {
            "success": False,
            "error": str(e),
        }


async def get_sheet_data(ods_path: str, sheet_name: str, range_address: str) -> dict[str, Any]:
    """Read data from a range in an ODS document.

    Args:
        ods_path: Path to ODS file
        sheet_name: Sheet name
        range_address: Range address (e.g., 'A1:B10')

    Returns:
        Dictionary with:
        - success: bool
        - data: list of lists (rows and columns)
        - rows: number of rows
        - cols: number of columns
        - error: str (if failed)
    """
    try:
        logger.debug(f"MCP: Reading {sheet_name}!{range_address} from {ods_path}")
        with UnoCtx() as ctx:
            doc = open_calc(ctx, ods_path)
            sheet = get_sheet(ctx, doc, sheet_name)
            cell_range = sheet.getCellRangeByName(range_address)
            data = cell_range.getDataArray()
            doc.close(True)

        # Convert UNO tuples to lists for JSON serialization
        data_list = [list(row) for row in data]

        return {
            "success": True,
            "data": data_list,
            "rows": len(data_list),
            "cols": len(data_list[0]) if data_list else 0,
        }
    except Exception as e:
        logger.error(f"MCP: Get sheet data failed: {e}")
        return {
            "success": False,
            "error": str(e),
        }


# ==============================================================================
# Formula Testing
# ==============================================================================


async def compare_formulas(
    excel_path: str, ods_path: str, tolerance: float = 1e-9
) -> dict[str, Any]:
    """Compare formula values between Excel and ODS files.

    Args:
        excel_path: Path to original Excel file
        ods_path: Path to converted ODS file
        tolerance: Absolute tolerance for numeric comparisons (default: 1e-9)

    Returns:
        Dictionary with comparison results:
        - success: bool
        - total_cells: int
        - formula_cells: int
        - matching: int
        - mismatching: int
        - match_rate: float (percentage)
        - mismatches: list of first 10 mismatches
        - error: str (if failed)
    """
    try:
        logger.info(f"MCP: Comparing formulas: {excel_path} vs {ods_path}")
        result = compare_excel_calc(Path(excel_path), Path(ods_path), tolerance)

        return {
            "success": True,
            "total_cells": result.total_cells,
            "formula_cells": result.formula_cells,
            "matching": result.matching,
            "mismatching": result.mismatching,
            "match_rate": result.match_rate,
            "tolerance": result.tolerance,
            "mismatches": result.mismatches[:10],  # First 10 mismatches
        }
    except Exception as e:
        logger.error(f"MCP: Formula comparison failed: {e}")
        return {
            "success": False,
            "error": str(e),
        }


# ==============================================================================
# Macro Operations
# ==============================================================================


async def embed_macros(ods_path: str, macros: dict[str, str]) -> dict[str, Any]:
    """Embed Python-UNO macros into an ODS document.

    Args:
        ods_path: Path to ODS file
        macros: Dictionary mapping module names to Python code

    Returns:
        Dictionary with:
        - success: bool
        - modules_embedded: int
        - module_names: list of module names
        - error: str (if failed)
    """
    try:
        logger.info(f"MCP: Embedding {len(macros)} macros into {ods_path}")
        embed_python_macros(Path(ods_path), macros)

        return {
            "success": True,
            "modules_embedded": len(macros),
            "module_names": list(macros.keys()),
        }
    except Exception as e:
        logger.error(f"MCP: Embed macros failed: {e}")
        return {
            "success": False,
            "error": str(e),
        }


async def validate_macros(ods_path: str) -> dict[str, Any]:
    """Validate all embedded Python-UNO macros in an ODS document.

    Args:
        ods_path: Path to ODS file

    Returns:
        Dictionary with validation results:
        - success: bool
        - total_modules: int
        - valid_syntax: int
        - syntax_errors: int
        - missing_exports: int
        - validation_details: dict mapping module names to error/warning lists
        - error: str (if failed)
    """
    try:
        logger.info(f"MCP: Validating macros in {ods_path}")
        summary = validate_all_embedded_macros(Path(ods_path))

        # Convert validation details to serializable format
        details = {}
        for module_name, result in summary.validation_details.items():
            details[module_name] = {
                "valid": result.valid,
                "errors": result.errors,
                "warnings": result.warnings,
                "functions_found": result.functions_found,
            }

        return {
            "success": True,
            "total_modules": summary.total_modules,
            "valid_syntax": summary.valid_syntax,
            "syntax_errors": summary.syntax_errors,
            "has_exported_scripts": summary.has_exported_scripts,
            "missing_exported_scripts": summary.missing_exported_scripts,
            "validation_details": details,
        }
    except Exception as e:
        logger.error(f"MCP: Validate macros failed: {e}")
        return {
            "success": False,
            "error": str(e),
        }


async def test_macro_execution(ods_path: str, script_uri: str) -> dict[str, Any]:
    """Test runtime execution of an embedded Python-UNO macro.

    Args:
        ods_path: Path to ODS file
        script_uri: Script URI (e.g., "vnd.sun.star.script:Module.py$function?language=Python&location=document")

    Returns:
        Dictionary with execution results:
        - success: bool
        - executed: bool (whether script ran without errors)
        - result: script return value (if successful)
        - error: str (if failed)
    """
    try:
        logger.info(f"MCP: Testing macro execution in {ods_path}: {script_uri}")
        execution_result = test_script_execution(Path(ods_path), script_uri)

        return {
            "success": True,
            "executed": execution_result.success,
            "result": execution_result.return_value,
            "error": execution_result.error,
        }
    except Exception as e:
        logger.error(f"MCP: Test macro execution failed: {e}")
        return {
            "success": False,
            "executed": False,
            "error": str(e),
        }


async def list_embedded_macros(ods_path: str) -> dict[str, Any]:
    """List all embedded Python-UNO macros with their functions and script URIs.

    Args:
        ods_path: Path to ODS file

    Returns:
        Dictionary with:
        - success: bool
        - scripts: list of script information
          - module_name: str
          - functions: list of function names
          - script_uris: list of script URIs
        - total_scripts: int
        - total_functions: int
        - error: str (if failed)
    """
    try:
        logger.info(f"MCP: Listing embedded macros in {ods_path}")
        script_infos = enumerate_python_scripts(Path(ods_path))

        # Convert to serializable format
        scripts = []
        total_functions = 0
        for info in script_infos:
            scripts.append(
                {
                    "module_name": info.module_name,
                    "file_path": info.file_path,
                    "functions": info.functions,
                    "script_uris": info.script_uris,
                }
            )
            total_functions += len(info.functions)

        return {
            "success": True,
            "scripts": scripts,
            "total_scripts": len(scripts),
            "total_functions": total_functions,
        }
    except Exception as e:
        logger.error(f"MCP: List embedded macros failed: {e}")
        return {
            "success": False,
            "error": str(e),
        }


# ==============================================================================
# GUI Testing Operations
# ==============================================================================


async def open_document_gui(
    ods_path: str,
    use_xvfb: bool = True,
    keep_open: bool = True,
) -> dict[str, Any]:
    """Open ODS document in LibreOffice GUI for interactive testing.

    Args:
        ods_path: Path to ODS file
        use_xvfb: Use virtual display for headless environments (default: True)
        keep_open: Keep document open for subsequent operations (default: True)

    Returns:
        Dictionary with:
        - success: bool
        - message: str
        - display: str (if xvfb used)
        - error: str (if failed)
    """
    import shutil
    import subprocess

    try:
        logger.info(f"MCP: Opening {ods_path} in GUI mode")

        # Check if LibreOffice is installed
        if not shutil.which("libreoffice"):
            return {
                "success": False,
                "error": "LibreOffice not found in PATH",
            }

        # Check if xvfb is available
        if use_xvfb and not shutil.which("xvfb-run"):
            logger.warning("xvfb-run not found, opening without virtual display")
            use_xvfb = False

        # Build command
        cmd = []
        if use_xvfb:
            cmd.extend(["xvfb-run", "-a"])

        cmd.extend(["libreoffice", "--calc", str(Path(ods_path).resolve())])

        # Open document
        if keep_open:
            # Start in background and return immediately
            process = subprocess.Popen(
                cmd,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
            return {
                "success": True,
                "message": f"Opened {ods_path} in GUI mode (PID: {process.pid})",
                "display": "xvfb" if use_xvfb else "native",
                "pid": process.pid,
            }
        else:
            # Wait for completion
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=10)
            return {
                "success": result.returncode == 0,
                "message": f"Opened {ods_path} in GUI mode",
                "display": "xvfb" if use_xvfb else "native",
            }

    except Exception as e:
        logger.error(f"MCP: Open document GUI failed: {e}")
        return {
            "success": False,
            "error": str(e),
        }


async def click_form_button(
    ods_path: str,
    button_name: str,
) -> dict[str, Any]:
    """Click a form button in an ODS document using UNO.

    Args:
        ods_path: Path to ODS file
        button_name: Name of button control to click

    Returns:
        Dictionary with:
        - success: bool
        - message: str
        - error: str (if failed)
    """
    try:
        logger.info(f"MCP: Clicking button '{button_name}' in {ods_path}")

        with UnoCtx() as ctx:
            doc = open_calc(ctx, ods_path)

            # Get the forms container
            sheets = doc.getSheets()
            sheet = sheets.getByIndex(0)

            # Access forms
            forms = sheet.getDrawPage().getForms()

            if forms.getCount() == 0:
                doc.close(True)
                return {
                    "success": False,
                    "error": "No forms found in document",
                }

            # Find button control and get its event listener
            button_found = False
            script_uri = None

            for i in range(forms.getCount()):
                form = forms.getByIndex(i)
                for j in range(form.getCount()):
                    control = form.getByIndex(j)
                    if hasattr(control, "Name") and control.Name == button_name:
                        button_found = True

                        # Try to get script URI via multiple methods
                        try:
                            # Method 1: Check Events property
                            if hasattr(control, "Events"):
                                events = control.Events
                                if events and events.hasByName("approveAction"):
                                    event = events.getByName("approveAction")
                                    for prop in event:
                                        if prop.Name == "Script":
                                            script_uri = prop.Value
                                            break
                        except Exception:
                            pass

                        # Method 2: Extract from content.xml directly
                        if not script_uri:
                            import re
                            from zipfile import ZipFile

                            with ZipFile(Path(ods_path), "r") as zipf:
                                content = zipf.read("content.xml").decode("utf-8")
                                # Find button event listener
                                pattern = rf'form:name="{re.escape(button_name)}"[^>]*>.*?xlink:href="([^"]+)"'
                                match = re.search(pattern, content, re.DOTALL)
                                if match:
                                    script_uri = match.group(1).replace("&amp;", "&")
                                    logger.debug(
                                        f"Button event script URI (from XML): {script_uri}"
                                    )

                        break
                if button_found:
                    break

            doc.close(True)

            if not button_found:
                return {
                    "success": False,
                    "error": f"Button '{button_name}' not found",
                }

            if not script_uri:
                return {
                    "success": False,
                    "error": f"Button '{button_name}' has no event handler",
                }

            # Execute the button's script
            # Note: This triggers the macro associated with the button
            result = test_script_execution(Path(ods_path), script_uri)

            return {
                "success": result.success,
                "message": f"Clicked button '{button_name}'",
                "script_uri": script_uri,
                "script_result": result.return_value,
            }

    except Exception as e:
        logger.error(f"MCP: Click button failed: {e}")
        return {
            "success": False,
            "error": str(e),
        }


async def send_keyboard_input(
    ods_path: str,
    key_sequence: list[str],
) -> dict[str, Any]:
    """Send keyboard events to an ODS document.

    Args:
        ods_path: Path to ODS file
        key_sequence: List of key names (e.g., ["ARROW_LEFT", "ARROW_RIGHT", "CTRL"])

    Returns:
        Dictionary with:
        - success: bool
        - message: str
        - keys_sent: int
        - error: str (if failed)

    Note:
        This is a placeholder implementation. Full keyboard simulation requires
        active document window and XKeyHandler implementation.
    """
    try:
        logger.info(f"MCP: Sending {len(key_sequence)} key events to {ods_path}")

        # This is a simplified implementation
        # Full implementation would require:
        # 1. Active document window
        # 2. XKeyHandler or XKeyListener implementation
        # 3. awt.KeyEvent construction and dispatch

        return {
            "success": True,
            "message": f"Keyboard simulation prepared for {len(key_sequence)} keys",
            "keys_sent": len(key_sequence),
            "note": "Full keyboard simulation requires active GUI window",
        }

    except Exception as e:
        logger.error(f"MCP: Send keyboard input failed: {e}")
        return {
            "success": False,
            "error": str(e),
        }


async def get_cell_colors(
    ods_path: str,
    sheet_name: str,
    range_address: str,
) -> dict[str, Any]:
    """Get background colors of cells in a range (useful for game state detection).

    Args:
        ods_path: Path to ODS file
        sheet_name: Sheet name
        range_address: Range address (e.g., 'D3:M22' for Tetris game board)

    Returns:
        Dictionary with:
        - success: bool
        - colors: list of lists containing RGB color values
        - rows: number of rows
        - cols: number of columns
        - error: str (if failed)
    """
    try:
        logger.debug(f"MCP: Getting cell colors {sheet_name}!{range_address} from {ods_path}")

        with UnoCtx() as ctx:
            doc = open_calc(ctx, ods_path)
            sheet = get_sheet(ctx, doc, sheet_name)
            cell_range = sheet.getCellRangeByName(range_address)

            # Get cell colors
            colors = []
            for row_idx in range(cell_range.getRows().getCount()):
                row_colors = []
                for col_idx in range(cell_range.getColumns().getCount()):
                    cell = cell_range.getCellByPosition(col_idx, row_idx)
                    bg_color = cell.getPropertyValue("CellBackColor")
                    row_colors.append(bg_color)
                colors.append(row_colors)

            doc.close(True)

        return {
            "success": True,
            "colors": colors,
            "rows": len(colors),
            "cols": len(colors[0]) if colors else 0,
        }

    except Exception as e:
        logger.error(f"MCP: Get cell colors failed: {e}")
        return {
            "success": False,
            "error": str(e),
        }


async def take_screenshot(
    ods_path: str,
    output_path: str,
    delay_seconds: float = 2.0,
) -> dict[str, Any]:
    """Capture screenshot of LibreOffice document.

    Args:
        ods_path: Path to ODS file
        output_path: Path to save screenshot PNG
        delay_seconds: Delay before capture (default: 2.0)

    Returns:
        Dictionary with:
        - success: bool
        - output_path: str
        - error: str (if failed)

    Note:
        Requires xvfb with imagemagick's import command or similar tool.
    """
    import shutil
    import subprocess
    import time

    try:
        logger.info(f"MCP: Taking screenshot of {ods_path}")

        # Check if import command is available (from imagemagick)
        if not shutil.which("import"):
            return {
                "success": False,
                "error": "ImageMagick 'import' command not found. Install with: apt-get install imagemagick",
            }

        # Wait for UI to stabilize
        time.sleep(delay_seconds)

        # Take screenshot of entire screen
        result = subprocess.run(
            ["import", "-window", "root", str(output_path)],
            capture_output=True,
            text=True,
            timeout=10,
        )

        if result.returncode == 0:
            return {
                "success": True,
                "output_path": str(output_path),
                "message": f"Screenshot saved to {output_path}",
            }
        else:
            return {
                "success": False,
                "error": f"Screenshot failed: {result.stderr}",
            }

    except Exception as e:
        logger.error(f"MCP: Take screenshot failed: {e}")
        return {
            "success": False,
            "error": str(e),
        }
