"""Direct Python execution validator using real UNO connection.

This module validates embedded Python-UNO macros by executing them using
a real LibreOffice UNO connection. Uses xvfb for GUI mode on headless servers.
"""

from dataclasses import dataclass
from pathlib import Path

from loguru import logger


@dataclass
class DirectExecutionResult:
    """Result of direct Python execution."""

    success: bool
    error: str | None
    output: str | None


def execute_script_directly(
    ods_path: Path, module_name: str, function_name: str
) -> DirectExecutionResult:
    """Execute a Python script using real UNO connection.

    Args:
        ods_path: Path to ODS file
        module_name: Module name (e.g., "Module1.py")
        function_name: Function name to execute

    Returns:
        DirectExecutionResult with execution details
    """
    try:
        from xlsliberator.python_macro_manager import test_script_execution

        # Build script URI
        script_uri = (
            f"vnd.sun.star.script:{module_name}${function_name}?language=Python&location=document"
        )

        # Use real UNO execution via XScriptProvider
        result = test_script_execution(ods_path, script_uri)

        return DirectExecutionResult(
            success=result.success,
            error=result.error,
            output=str(result.return_value) if result.return_value else None,
        )

    except Exception as e:
        return DirectExecutionResult(False, f"Execution failed: {e}", None)


def test_script_basic_execution(ods_path: Path, script_uri: str) -> DirectExecutionResult:
    """Test basic execution of a script using real UNO.

    Args:
        ods_path: Path to ODS file
        script_uri: Script URI (e.g., "vnd.sun.star.script:Module.py$function?...")

    Returns:
        DirectExecutionResult
    """
    try:
        # Parse script URI
        uri_parts = script_uri.split("vnd.sun.star.script:")[1]
        module_and_func = uri_parts.split("?")[0]
        module_name, function_name = module_and_func.split("$")

        logger.debug(f"Testing execution via UNO: {module_name} -> {function_name}")

        return execute_script_directly(ods_path, module_name, function_name)

    except Exception as e:
        return DirectExecutionResult(False, f"URI parsing failed: {e}", None)


def validate_all_scripts_direct(ods_path: Path) -> dict[str, DirectExecutionResult]:
    """Validate all scripts using real UNO connection.

    Args:
        ods_path: Path to ODS file

    Returns:
        Dictionary mapping script URIs to execution results
    """
    from xlsliberator.python_macro_manager import enumerate_python_scripts

    results = {}

    try:
        script_infos = enumerate_python_scripts(ods_path)

        for script_info in script_infos:
            for uri in script_info.script_uris:
                result = test_script_basic_execution(ods_path, uri)
                results[uri] = result
                if result.success:
                    logger.debug(f"✓ {uri}: Executed successfully")
                else:
                    # Check if it's expected XScriptProvider limitation
                    if result.error and "XScriptProvider" in result.error:
                        logger.warning(f"⚠ {uri}: XScriptProvider unavailable (GUI mode needed)")
                    else:
                        logger.warning(f"✗ {uri}: {result.error}")

    except Exception as e:
        logger.error(f"Failed to validate scripts: {e}")

    return results
