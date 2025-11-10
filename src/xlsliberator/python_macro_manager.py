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
    from xlsliberator.uno_conn import UnoCtx, open_calc

    ods_path = Path(ods_path)
    if not ods_path.exists():
        return ScriptExecutionResult(False, f"File not found: {ods_path}", None)

    try:
        with UnoCtx() as ctx:
            # Open document
            doc = open_calc(ctx, ods_path)

            try:
                # Get MasterScriptProvider
                msp_factory = ctx.component_context.ServiceManager.createInstanceWithContext(
                    "com.sun.star.script.provider.MasterScriptProviderFactory",
                    ctx.component_context,
                )
                script_provider = msp_factory.createScriptProvider(doc)

                # Get script
                script = script_provider.getScript(script_uri)

                # Execute script (with empty parameters)
                result = script.invoke((), (), ())

                logger.debug(f"Script executed successfully: {script_uri}")
                return ScriptExecutionResult(True, None, result)

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
