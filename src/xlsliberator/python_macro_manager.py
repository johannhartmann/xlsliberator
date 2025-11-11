"""Python macro management utilities using LibreOffice UNO API.

This module provides utilities for:
1. Script Validation - Validate Python syntax and UNO compatibility
2. Script Execution Testing - Test embedded scripts via UNO ScriptProvider
3. Script Enumeration - List Python scripts in ODS files
4. Script Editing - Edit embedded scripts (via embed_macros.py)
"""

import ast
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from loguru import logger


class PythonMacroError(Exception):
    """Raised when Python macro operations fail."""


@dataclass
class ScriptInfo:
    """Information about an embedded Python script."""

    module_name: str
    file_path: str
    functions: list[str]
    script_uris: list[str]


@dataclass
class ScriptValidationResult:
    """Result of script validation."""

    valid: bool
    errors: list[str]
    warnings: list[str]
    functions_found: list[str]


@dataclass
class ScriptExecutionResult:
    """Result of script execution test."""

    success: bool
    error: str | None
    return_value: Any | None


# ============================================================================
# 1. Script Validation
# ============================================================================


def validate_python_script(script_code: str) -> ScriptValidationResult:
    """Validate Python script syntax and structure.

    Args:
        script_code: Python source code to validate

    Returns:
        ScriptValidationResult with validation details
    """
    errors: list[str] = []
    warnings: list[str] = []
    functions_found: list[str] = []

    # Check for empty code
    if not script_code.strip():
        errors.append("Script is empty")
        return ScriptValidationResult(False, errors, warnings, functions_found)

    # Parse AST
    try:
        tree = ast.parse(script_code)
    except SyntaxError as e:
        errors.append(f"Syntax error at line {e.lineno}: {e.msg}")
        return ScriptValidationResult(False, errors, warnings, functions_found)
    except Exception as e:
        errors.append(f"Failed to parse script: {e}")
        return ScriptValidationResult(False, errors, warnings, functions_found)

    # Analyze AST
    has_uno_import = False
    has_xscriptcontext = False

    for node in ast.walk(tree):
        # Find function definitions
        if isinstance(node, ast.FunctionDef):
            functions_found.append(node.name)

        # Check for UNO imports
        elif isinstance(node, ast.Import):
            for alias in node.names:
                if alias.name == "uno":
                    has_uno_import = True

        # Check for XSCRIPTCONTEXT usage
        elif isinstance(node, ast.Name) and node.id == "XSCRIPTCONTEXT":
            has_xscriptcontext = True

    # Validation checks
    if not functions_found:
        warnings.append("No functions found in script")

    if not has_uno_import and has_xscriptcontext:
        warnings.append("Uses XSCRIPTCONTEXT but doesn't import uno")

    # Check for g_exportedScripts
    has_exported = False
    for node in tree.body:
        if isinstance(node, ast.Assign):
            for target in node.targets:
                if isinstance(target, ast.Name) and target.id == "g_exportedScripts":
                    has_exported = True
                    break

    if not has_exported and functions_found:
        warnings.append(
            "No g_exportedScripts defined - functions won't be callable from LibreOffice"
        )

    # Try to compile
    try:
        compile(script_code, "<string>", "exec")
    except SyntaxError as e:
        errors.append(f"Compilation error at line {e.lineno}: {e.msg}")
    except Exception as e:
        errors.append(f"Compilation failed: {e}")

    valid = len(errors) == 0
    return ScriptValidationResult(valid, errors, warnings, functions_found)


# ============================================================================
# 2. Script Execution Testing
# ============================================================================


def test_script_execution(ods_path: str | Path, script_uri: str) -> ScriptExecutionResult:
    """Test execution of an embedded Python script.

    Args:
        ods_path: Path to ODS file with embedded scripts
        script_uri: Script URI (e.g., "vnd.sun.star.script:Module.py$function?language=Python&location=document")

    Returns:
        ScriptExecutionResult with execution details
    """
    import signal

    from xlsliberator.uno_conn import UnoCtx, open_calc

    ods_path = Path(ods_path)
    if not ods_path.exists():
        return ScriptExecutionResult(False, f"File not found: {ods_path}", None)

    def timeout_handler(_signum: int, _frame: Any) -> None:
        raise TimeoutError("Script execution timed out")

    try:
        with UnoCtx(use_gui=True) as ctx:
            # Open document
            doc = open_calc(ctx, ods_path)

            try:
                # Set timeout for script execution (5 seconds)
                signal.signal(signal.SIGALRM, timeout_handler)
                signal.alarm(5)

                try:
                    # Get MasterScriptProvider
                    msp_factory = ctx.component_context.ServiceManager.createInstanceWithContext(
                        "com.sun.star.script.provider.MasterScriptProviderFactory",
                        ctx.component_context,
                    )
                    script_provider = msp_factory.createScriptProvider(doc)

                    # Get script (this may hang if XScriptProvider unavailable)
                    script = script_provider.getScript(script_uri)

                    # Execute script (with empty parameters)
                    result = script.invoke((), (), ())

                    logger.debug(f"Script executed successfully: {script_uri}")
                    return ScriptExecutionResult(True, None, result)

                finally:
                    # Cancel alarm
                    signal.alarm(0)

            except TimeoutError:
                error_msg = "Script execution timed out (XScriptProvider likely unavailable)"
                logger.warning(error_msg)
                return ScriptExecutionResult(False, error_msg, None)

            except Exception as e:
                error_msg = f"Script execution failed: {e}"
                logger.warning(error_msg)
                return ScriptExecutionResult(False, error_msg, None)

            finally:
                doc.close(True)

    except Exception as e:
        error_msg = f"Failed to open document: {e}"
        logger.error(error_msg)
        return ScriptExecutionResult(False, error_msg, None)


def test_all_scripts(ods_path: str | Path) -> dict[str, ScriptExecutionResult]:
    """Test execution of all embedded Python scripts.

    Args:
        ods_path: Path to ODS file with embedded scripts

    Returns:
        Dictionary mapping script URIs to execution results
    """
    results = {}

    # Get all script infos
    script_infos = enumerate_python_scripts(ods_path)

    # Test each script URI
    for script_info in script_infos:
        for uri in script_info.script_uris:
            result = test_script_execution(ods_path, uri)
            results[uri] = result

    return results


# ============================================================================
# 3. Script Enumeration
# ============================================================================


def enumerate_python_scripts(ods_path: str | Path) -> list[ScriptInfo]:
    """Enumerate all embedded Python scripts in an ODS file.

    Args:
        ods_path: Path to ODS file

    Returns:
        List of ScriptInfo objects describing each script
    """
    ods_path = Path(ods_path)
    if not ods_path.exists():
        raise PythonMacroError(f"File not found: {ods_path}")

    script_infos = []

    try:
        with zipfile.ZipFile(ods_path, "r") as zip_file:
            # Find all Python scripts
            for filename in zip_file.namelist():
                if filename.startswith("Scripts/python/") and filename.endswith(".py"):
                    # Read script content
                    script_code = zip_file.read(filename).decode("utf-8")

                    # Parse to find functions
                    functions = []
                    try:
                        tree = ast.parse(script_code)
                        for node in ast.walk(tree):
                            if isinstance(node, ast.FunctionDef):
                                functions.append(node.name)
                    except Exception as e:
                        logger.warning(f"Failed to parse {filename} for function extraction: {e}")

                    # Generate script URIs for each function
                    module_name = Path(filename).name
                    script_uris = [
                        f"vnd.sun.star.script:{module_name}${func}?language=Python&location=document"
                        for func in functions
                    ]

                    script_info = ScriptInfo(
                        module_name=module_name,
                        file_path=filename,
                        functions=functions,
                        script_uris=script_uris,
                    )
                    script_infos.append(script_info)

        logger.info(f"Found {len(script_infos)} Python scripts in {ods_path}")
        return script_infos

    except zipfile.BadZipFile as e:
        raise PythonMacroError(f"Invalid ODS file: {e}") from e
    except Exception as e:
        raise PythonMacroError(f"Failed to enumerate scripts: {e}") from e


def list_python_scripts(ods_path: str | Path) -> list[str]:
    """List all embedded Python script module names.

    Args:
        ods_path: Path to ODS file

    Returns:
        List of script module names (e.g., ["Game.bas.py", "Engine.bas.py"])
    """
    script_infos = enumerate_python_scripts(ods_path)
    return [info.module_name for info in script_infos]


def get_script_functions(ods_path: str | Path, module_name: str) -> list[str]:
    """Get all functions in a specific Python script module.

    Args:
        ods_path: Path to ODS file
        module_name: Name of Python module (e.g., "Game.bas.py")

    Returns:
        List of function names
    """
    script_infos = enumerate_python_scripts(ods_path)

    for info in script_infos:
        if info.module_name == module_name:
            return info.functions

    raise PythonMacroError(f"Module not found: {module_name}")


def generate_script_uri(module_name: str, function_name: str) -> str:
    """Generate a script URI for a Python function.

    Args:
        module_name: Python module name (e.g., "Game.bas.py")
        function_name: Function name (e.g., "StartButton_Click")

    Returns:
        Script URI string
    """
    return f"vnd.sun.star.script:{module_name}${function_name}?language=Python&location=document"


# ============================================================================
# 4. Script Editing
# ============================================================================

# Script editing is already implemented in embed_macros.py
# No additional implementation needed here


# ============================================================================
# 5. Post-Conversion Validation
# ============================================================================


@dataclass
class MacroValidationSummary:
    """Summary of macro validation results."""

    total_modules: int
    valid_syntax: int
    syntax_errors: int
    has_exported_scripts: int
    missing_exported_scripts: int
    validation_details: dict[str, ScriptValidationResult]


def validate_all_embedded_macros(ods_path: str | Path) -> MacroValidationSummary:
    """Validate all embedded Python macros in an ODS file.

    Args:
        ods_path: Path to ODS file

    Returns:
        MacroValidationSummary with comprehensive validation results
    """
    import zipfile

    ods_path = Path(ods_path)
    if not ods_path.exists():
        raise PythonMacroError(f"File not found: {ods_path}")

    validation_details = {}
    total_modules = 0
    valid_syntax = 0
    syntax_errors = 0
    has_exported_scripts = 0
    missing_exported_scripts = 0

    try:
        with zipfile.ZipFile(ods_path, "r") as zip_file:
            # Find all Python scripts
            for filename in zip_file.namelist():
                if filename.startswith("Scripts/python/") and filename.endswith(".py"):
                    total_modules += 1
                    module_name = Path(filename).name

                    # Read script content
                    script_code = zip_file.read(filename).decode("utf-8")

                    # Validate
                    result = validate_python_script(script_code)
                    validation_details[module_name] = result

                    # Update counters
                    if result.valid:
                        valid_syntax += 1
                    else:
                        syntax_errors += 1

                    # Check for g_exportedScripts
                    if "g_exportedScripts" in script_code:
                        has_exported_scripts += 1
                    else:
                        missing_exported_scripts += 1

        logger.info(
            f"Validated {total_modules} Python modules: "
            f"{valid_syntax} valid, {syntax_errors} with errors"
        )

        return MacroValidationSummary(
            total_modules=total_modules,
            valid_syntax=valid_syntax,
            syntax_errors=syntax_errors,
            has_exported_scripts=has_exported_scripts,
            missing_exported_scripts=missing_exported_scripts,
            validation_details=validation_details,
        )

    except zipfile.BadZipFile as e:
        raise PythonMacroError(f"Invalid ODS file: {e}") from e
    except Exception as e:
        raise PythonMacroError(f"Failed to validate macros: {e}") from e


# ============================================================================
# 6. Macro Execution Testing with Undo Contexts
# ============================================================================


@dataclass
class MacroExecutionSummary:
    """Summary of macro execution testing results."""

    total_functions: int
    successful: int
    failed: int
    skipped: int
    execution_details: dict[str, ScriptExecutionResult]


def test_all_macros_safe(ods_path: str | Path) -> MacroExecutionSummary:
    """Test all embedded Python macros with Undo context safety.

    Creates a temporary copy of the document and tests all macros with
    automatic rollback on errors.

    Args:
        ods_path: Path to ODS file

    Returns:
        MacroExecutionSummary with test results
    """
    import shutil
    import tempfile

    ods_path = Path(ods_path)
    if not ods_path.exists():
        raise PythonMacroError(f"File not found: {ods_path}")

    execution_details = {}
    total_functions = 0
    successful = 0
    failed = 0
    skipped = 0

    # Create temporary copy for safe testing
    with tempfile.TemporaryDirectory() as tmpdir:
        test_ods = Path(tmpdir) / "test.ods"
        shutil.copy(ods_path, test_ods)

        try:
            # Get all script infos
            script_infos = enumerate_python_scripts(test_ods)

            for script_info in script_infos:
                for function_name in script_info.functions:
                    total_functions += 1

                    # Skip private functions and special methods
                    if function_name.startswith("_"):
                        skipped += 1
                        logger.debug(f"Skipping private function: {function_name}")
                        continue

                    # Generate script URI
                    script_uri = generate_script_uri(script_info.module_name, function_name)

                    # Test execution
                    logger.debug(f"Testing macro: {script_uri}")
                    result = test_script_execution_safe(test_ods, script_uri)

                    execution_details[script_uri] = result

                    if result.success:
                        successful += 1
                        logger.debug(f"✓ {function_name} executed successfully")
                    else:
                        failed += 1
                        logger.warning(f"✗ {function_name} failed: {result.error}")

            logger.info(
                f"Tested {total_functions} functions: "
                f"{successful} passed, {failed} failed, {skipped} skipped"
            )

        except Exception as e:
            logger.error(f"Macro testing failed: {e}")
            raise PythonMacroError(f"Failed to test macros: {e}") from e

    return MacroExecutionSummary(
        total_functions=total_functions,
        successful=successful,
        failed=failed,
        skipped=skipped,
        execution_details=execution_details,
    )


def test_script_execution_safe(ods_path: str | Path, script_uri: str) -> ScriptExecutionResult:
    """Test execution of a script with Undo context for safety.

    Args:
        ods_path: Path to ODS file (should be a test copy)
        script_uri: Script URI to execute

    Returns:
        ScriptExecutionResult
    """
    import contextlib
    import signal

    from xlsliberator.uno_conn import UnoCtx, open_calc

    def timeout_handler(_signum: int, _frame: Any) -> None:
        raise TimeoutError("Script execution timed out")

    try:
        with UnoCtx(use_gui=True) as ctx:
            doc = open_calc(ctx, ods_path)

            try:
                # Get undo manager
                um = doc.getUndoManager()
                um.enterUndoContext("Test Macro Execution")

                # Set timeout for script execution (5 seconds)
                signal.signal(signal.SIGALRM, timeout_handler)
                signal.alarm(5)

                try:
                    # Get script provider
                    msp_factory = ctx.component_context.ServiceManager.createInstanceWithContext(
                        "com.sun.star.script.provider.MasterScriptProviderFactory",
                        ctx.component_context,
                    )
                    script_provider = msp_factory.createScriptProvider(doc)

                    # Get and execute script (this may hang if XScriptProvider unavailable)
                    script = script_provider.getScript(script_uri)
                    result = script.invoke((), (), ())

                    # Success - leave undo context
                    um.leaveUndoContext()

                    logger.debug(f"Script executed successfully: {script_uri}")
                    return ScriptExecutionResult(True, None, result)

                except TimeoutError:
                    error_msg = "Script execution timed out (XScriptProvider likely unavailable)"
                    logger.warning(error_msg)
                    # Try to leave undo context, ignore if it fails
                    with contextlib.suppress(Exception):
                        um.leaveUndoContext()
                    return ScriptExecutionResult(False, error_msg, None)

                except Exception as e:
                    # Execution failed - undo changes
                    error_msg = f"Script execution failed: {e}"
                    logger.debug(error_msg)

                    try:
                        um.leaveUndoContext()
                        um.undo()
                    except Exception as undo_error:
                        logger.warning(f"Undo failed: {undo_error}")

                    return ScriptExecutionResult(False, error_msg, None)

                finally:
                    # Cancel alarm
                    signal.alarm(0)

            finally:
                doc.close(False)  # Don't save

    except Exception as e:
        error_msg = f"Failed to open document: {e}"
        logger.error(error_msg)
        return ScriptExecutionResult(False, error_msg, None)


# ============================================================================
# 7. Formula Validation with FunctionAccess
# ============================================================================


@dataclass
class FormulaValidationResult:
    """Result of formula validation."""

    valid: bool
    error: str | None
    formula: str
    result: Any | None


@dataclass
class FormulaValidationSummary:
    """Summary of formula validation for a document."""

    total_formulas: int
    valid_formulas: int
    invalid_formulas: int
    validation_details: dict[str, FormulaValidationResult]


def validate_formula(formula: str) -> FormulaValidationResult:
    """Validate a formula using FunctionAccess service.

    Args:
        formula: Formula string (without leading "=")

    Returns:
        FormulaValidationResult with validation details

    Example:
        >>> result = validate_formula("SUM(1,2,3)")
        >>> result.valid
        True
        >>> result.result
        6.0
    """
    from xlsliberator.uno_conn import UnoCtx

    # Strip leading "=" if present
    formula = formula.lstrip("=")

    try:
        with UnoCtx() as ctx:
            # Create FunctionAccess service
            smgr = ctx.component_context.ServiceManager
            fa = smgr.createInstanceWithContext(
                "com.sun.star.sheet.FunctionAccess", ctx.component_context
            )

            try:
                # Try to evaluate the formula
                # Note: This is a simplified approach that works for simple function calls
                # More complex formulas with cell references may not work
                result = fa.callFunction(formula, ())

                logger.debug(f"Formula validated successfully: {formula} = {result}")
                return FormulaValidationResult(
                    valid=True, error=None, formula=formula, result=result
                )

            except Exception as e:
                error_msg = f"Formula evaluation failed: {e}"
                logger.debug(error_msg)
                return FormulaValidationResult(
                    valid=False, error=error_msg, formula=formula, result=None
                )

    except Exception as e:
        error_msg = f"Failed to create FunctionAccess service: {e}"
        logger.error(error_msg)
        return FormulaValidationResult(valid=False, error=error_msg, formula=formula, result=None)


def validate_all_formulas(ods_path: str | Path) -> FormulaValidationSummary:
    """Validate all formulas in an ODS file.

    Args:
        ods_path: Path to ODS file

    Returns:
        FormulaValidationSummary with validation results

    Note:
        This extracts formulas from all cells and validates them using FunctionAccess.
        Complex formulas with cell references may not be fully validated.
    """
    from xlsliberator.uno_conn import UnoCtx, open_calc

    ods_path = Path(ods_path)
    if not ods_path.exists():
        raise PythonMacroError(f"File not found: {ods_path}")

    validation_details = {}
    total_formulas = 0
    valid_formulas = 0
    invalid_formulas = 0

    try:
        with UnoCtx() as ctx:
            doc = open_calc(ctx, ods_path)

            try:
                # Iterate through all sheets
                sheets = doc.getSheets()
                for sheet_idx in range(sheets.getCount()):
                    sheet = sheets.getByIndex(sheet_idx)
                    sheet_name = sheet.getName()

                    # Get all formula cells
                    # CellFlags.FORMULA = 2 (com.sun.star.sheet.CellFlags.FORMULA)
                    formula_cells = sheet.queryContentCells(2)

                    # Iterate through formula cells
                    cell_ranges = formula_cells.getRangeAddresses()
                    for cell_range in cell_ranges:
                        for row in range(cell_range.StartRow, cell_range.EndRow + 1):
                            for col in range(cell_range.StartColumn, cell_range.EndColumn + 1):
                                cell = sheet.getCellByPosition(col, row)
                                formula = cell.getFormula()

                                if formula and formula.startswith("="):
                                    total_formulas += 1
                                    cell_ref = f"{sheet_name}!{_get_cell_address(row, col)}"

                                    # Validate formula
                                    result = validate_formula(formula)
                                    validation_details[cell_ref] = result

                                    if result.valid:
                                        valid_formulas += 1
                                    else:
                                        invalid_formulas += 1

                logger.info(
                    f"Validated {total_formulas} formulas: "
                    f"{valid_formulas} valid, {invalid_formulas} invalid"
                )

            finally:
                doc.close(True)

    except Exception as e:
        logger.error(f"Formula validation failed: {e}")
        raise PythonMacroError(f"Failed to validate formulas: {e}") from e

    return FormulaValidationSummary(
        total_formulas=total_formulas,
        valid_formulas=valid_formulas,
        invalid_formulas=invalid_formulas,
        validation_details=validation_details,
    )


def _get_cell_address(row: int, col: int) -> str:
    """Convert row/col indices to Excel-style address (A1, B2, etc).

    Args:
        row: Row index (0-based)
        col: Column index (0-based)

    Returns:
        Cell address string
    """
    # Convert column index to letter(s)
    col_str = ""
    col_num = col + 1
    while col_num > 0:
        col_num, remainder = divmod(col_num - 1, 26)
        col_str = chr(65 + remainder) + col_str

    return f"{col_str}{row + 1}"
