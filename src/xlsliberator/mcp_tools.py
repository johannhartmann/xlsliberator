"""MCP tool implementations for LibreOffice UNO operations.

Exposes xlsliberator functionality through Model Context Protocol (MCP) tools
for integration with Open-SWE and other explicitly configured MCP clients.
"""

from pathlib import Path
from typing import Any

from loguru import logger

from xlsliberator.api import convert as convert_api
from xlsliberator.boundary_models import (
    BoundaryError,
    BoundaryResponse,
    EvidenceRecord,
    RuntimeToolOptions,
)
from xlsliberator.docker_runtime import LibreOfficeDockerRuntime
from xlsliberator.embed_macros import embed_python_macros
from xlsliberator.execution_sandbox import WorkspacePathPolicy
from xlsliberator.python_macro_manager import (
    enumerate_python_scripts,
    validate_all_embedded_macros,
)
from xlsliberator.testing_lo import compare_excel_calc
from xlsliberator.validation_models import GateExecutionStatus
from xlsliberator.workbook_security import validate_untrusted_workbook


def _tool_response(
    status: GateExecutionStatus,
    *,
    data: dict[str, Any] | None = None,
    transport_success: bool = True,
    implemented: bool = True,
    capability_available: bool = True,
    evidence: list[EvidenceRecord] | None = None,
    error: BoundaryError | None = None,
) -> dict[str, Any]:
    """Build one canonical typed response and preserve legacy top-level data."""
    return BoundaryResponse(
        transport_success=transport_success,
        operation_status=status,
        implemented=implemented,
        capability_available=capability_available,
        evidence=evidence or [],
        error=error,
        data=data or {},
    ).to_payload()


def _operation_error(exc: BaseException) -> dict[str, Any]:
    """Return a truthful operation failure after a successful MCP call transport."""
    return _tool_response(
        GateExecutionStatus.FAILED,
        error=BoundaryError(type=type(exc).__name__, message=str(exc)),
    )


def _worker_tool_response(
    payload: dict[str, Any], runtime_options: RuntimeToolOptions | None = None
) -> dict[str, Any]:
    from xlsliberator.lo_worker_client import LibreOfficeWorkerClient, worker_unavailable_message

    options = runtime_options or RuntimeToolOptions()
    workspace = WorkspacePathPolicy()
    payload = dict(payload)
    for key in ("input_path", "ods_path", "excel_path"):
        if payload.get(key):
            payload[key] = str(workspace.input_file(str(payload[key])))
    if payload.get("output_path"):
        payload["output_path"] = str(workspace.output_file(str(payload["output_path"])))
    payload = {**payload, "timeout_seconds": options.timeout_seconds}
    runtime = LibreOfficeDockerRuntime(
        image=options.target_runtime_image,
        timeout_seconds=options.timeout_seconds,
        workspace_roots=list(workspace.roots),
    )
    response = LibreOfficeWorkerClient(
        timeout_seconds=options.timeout_seconds,
        runtime=runtime,
    ).request(payload)
    if response.success:
        return _tool_response(
            GateExecutionStatus.PASSED,
            data=response.data,
            evidence=[EvidenceRecord(kind="worker_response", data=response.to_dict())],
        )
    error_type = response.error.type if response.error else "WorkerError"
    transport_failure = error_type in {
        "DockerRuntimeUnavailable",
        "DockerRuntimeTimeout",
        "MalformedWorkerJSON",
        "MalformedWorkerResponse",
    }
    unavailable = error_type == "DockerRuntimeUnavailable"
    return _tool_response(
        GateExecutionStatus.UNAVAILABLE if unavailable else GateExecutionStatus.FAILED,
        transport_success=not transport_failure,
        capability_available=not unavailable,
        data={"worker_error": response.error.to_dict() if response.error else None},
        evidence=[EvidenceRecord(kind="worker_response", data=response.to_dict())],
        error=BoundaryError(type=error_type, message=worker_unavailable_message(response)),
    )


# ==============================================================================
# Document Operations
# ==============================================================================


async def convert_excel_to_ods(
    excel_path: str,
    output_path: str,
    embed_macros: bool = False,
) -> dict[str, Any]:
    """Convert Excel file to LibreOffice Calc ODS format.

    Args:
        excel_path: Path to input Excel file (.xlsx, .xlsm, .xlsb, .xls)
        output_path: Path for output ODS file
        embed_macros: Deprecated compatibility flag; no model translation is performed

    Returns:
        Dictionary with conversion results:
        - success: bool
        - output_path: str
        - report: dict with conversion statistics
        - error: str (if failed)
    """
    try:
        logger.info(f"MCP: Converting {excel_path} to {output_path}")
        workspace = WorkspacePathPolicy()
        source = workspace.input_file(excel_path)
        destination = workspace.output_file(output_path)
        validate_untrusted_workbook(source)
        report = convert_api(
            input_path=source,
            output_path=destination,
            embed_macros=embed_macros,
        )

        report_data = {
            "output_path": str(destination),
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
        return _tool_response(
            GateExecutionStatus.PASSED if report.success else GateExecutionStatus.FAILED,
            data=report_data,
            evidence=[
                EvidenceRecord(
                    kind="conversion_report",
                    data={"success": report.success, "errors": report.errors},
                )
            ],
            error=(
                BoundaryError(type="ConversionFailed", message="; ".join(report.errors))
                if report.errors
                else None
            ),
        )
    except Exception as e:
        logger.error(f"MCP: Conversion failed: {e}")
        return _operation_error(e)


async def inspect_workbook(excel_path: str) -> dict[str, Any]:
    """Inspect source workbook parse inventory."""
    try:
        logger.info(f"MCP: Inspecting workbook {excel_path}")
        from xlsliberator.inspect_workbook import (
            inspect_workbook as inspect_workbook_api,
        )
        from xlsliberator.inspect_workbook import (
            inventory_to_dict,
        )

        source = WorkspacePathPolicy().input_file(excel_path)
        validate_untrusted_workbook(source)
        inventory = inspect_workbook_api(source)
        inventory_data = inventory_to_dict(inventory)
        return _tool_response(
            GateExecutionStatus.PASSED,
            data={"inventory": inventory_data},
            evidence=[EvidenceRecord(kind="source_inventory", data=inventory_data)],
        )
    except Exception as e:
        logger.error(f"MCP: Inspect workbook failed: {e}")
        return _operation_error(e)


async def validate_transformation(
    excel_path: str,
    ods_path: str | None = None,
    target: str = "libreoffice",
) -> dict[str, Any]:
    """Run validation gates for a workbook transformation."""
    try:
        logger.info(f"MCP: Validating transformation {excel_path} -> {ods_path}")
        from xlsliberator.validation_runner import (
            ValidationPlan,
            ValidationRunner,
            parse_target_kind,
        )

        workspace = WorkspacePathPolicy()
        report = ValidationRunner(
            ValidationPlan(
                input_path=workspace.input_file(excel_path),
                output_path=workspace.input_file(ods_path) if ods_path else None,
                target_kinds=parse_target_kind(target),
            )
        ).run_all()
        certified = report.certification.certified
        certification = report.certification.model_dump(mode="json")
        return _tool_response(
            GateExecutionStatus.PASSED if certified else GateExecutionStatus.FAILED,
            data={"certification": certification},
            evidence=[EvidenceRecord(kind="certification", data=certification)],
            error=(
                None
                if certified
                else BoundaryError(
                    type="CertificationFailed",
                    message="Required validation gates did not all pass",
                )
            ),
        )
    except Exception as e:
        logger.error(f"MCP: Validate transformation failed: {e}")
        return _operation_error(e)


async def list_controls(ods_path: str) -> dict[str, Any]:
    """List discovered ODS form controls."""
    try:
        from xlsliberator.control_inventory import extract_controls_from_ods

        controls = extract_controls_from_ods(WorkspacePathPolicy().input_file(ods_path))
        serialized = [control.model_dump(mode="json") for control in controls]
        return _tool_response(
            GateExecutionStatus.PASSED,
            data={"controls": serialized, "count": len(controls)},
            evidence=[EvidenceRecord(kind="control_inventory", data={"controls": serialized})],
        )
    except Exception as e:
        logger.error(f"MCP: List controls failed: {e}")
        return _operation_error(e)


async def list_event_bindings(ods_path: str) -> dict[str, Any]:
    """List discovered ODS event bindings."""
    try:
        from xlsliberator.control_inventory import extract_event_bindings_from_ods

        event_bindings = extract_event_bindings_from_ods(WorkspacePathPolicy().input_file(ods_path))
        serialized = [event_binding.model_dump(mode="json") for event_binding in event_bindings]
        return _tool_response(
            GateExecutionStatus.PASSED,
            data={"event_bindings": serialized, "count": len(event_bindings)},
            evidence=[EvidenceRecord(kind="event_binding_inventory", data={"items": serialized})],
        )
    except Exception as e:
        logger.error(f"MCP: List event bindings failed: {e}")
        return _operation_error(e)


async def recalculate_document(
    ods_path: str, runtime_options: RuntimeToolOptions | None = None
) -> dict[str, Any]:
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
        result = _worker_tool_response(
            {"op": "recalculate_document", "ods_path": ods_path}, runtime_options
        )
        if not result["success"]:
            return result

        result["message"] = f"Recalculated {ods_path}"
        return result
    except Exception as e:
        logger.error(f"MCP: Recalculation failed: {e}")
        return _operation_error(e)


# ==============================================================================
# Cell and Sheet Operations
# ==============================================================================


async def read_cell(
    ods_path: str,
    sheet_name: str,
    cell_address: str,
    runtime_options: RuntimeToolOptions | None = None,
) -> dict[str, Any]:
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
        return _worker_tool_response(
            {
                "op": "read_cell",
                "ods_path": ods_path,
                "sheet_name": sheet_name,
                "cell_address": cell_address,
                "timeout_seconds": 30,
            },
            runtime_options,
        )
    except Exception as e:
        logger.error(f"MCP: Read cell failed: {e}")
        return _operation_error(e)


async def list_sheets(
    ods_path: str, runtime_options: RuntimeToolOptions | None = None
) -> dict[str, Any]:
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
        return _worker_tool_response({"op": "list_sheets", "ods_path": ods_path}, runtime_options)
    except Exception as e:
        logger.error(f"MCP: List sheets failed: {e}")
        return _operation_error(e)


async def get_sheet_data(
    ods_path: str,
    sheet_name: str,
    range_address: str,
    runtime_options: RuntimeToolOptions | None = None,
) -> dict[str, Any]:
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
        return _worker_tool_response(
            {
                "op": "get_sheet_data",
                "ods_path": ods_path,
                "sheet_name": sheet_name,
                "range_address": range_address,
                "timeout_seconds": 30,
            },
            runtime_options,
        )
    except Exception as e:
        logger.error(f"MCP: Get sheet data failed: {e}")
        return _operation_error(e)


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
        workspace = WorkspacePathPolicy()
        result = compare_excel_calc(
            workspace.input_file(excel_path), workspace.input_file(ods_path), tolerance
        )

        data = {
            "total_cells": result.total_cells,
            "formula_cells": result.formula_cells,
            "matching": result.matching,
            "mismatching": result.mismatching,
            "match_rate": result.match_rate,
            "tolerance": result.tolerance,
            "mismatches": result.mismatches[:10],  # First 10 mismatches
        }
        equivalent = result.mismatching == 0
        return _tool_response(
            GateExecutionStatus.PASSED if equivalent else GateExecutionStatus.FAILED,
            data=data,
            evidence=[EvidenceRecord(kind="formula_comparison", data=data)],
            error=(
                None
                if equivalent
                else BoundaryError(
                    type="FormulaMismatch",
                    message=f"{result.mismatching} formula result(s) differ",
                )
            ),
        )
    except Exception as e:
        logger.error(f"MCP: Formula comparison failed: {e}")
        return _operation_error(e)


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
        embed_python_macros(WorkspacePathPolicy().input_file(ods_path), macros)

        data = {"modules_embedded": len(macros), "module_names": list(macros)}
        return _tool_response(
            GateExecutionStatus.PASSED,
            data=data,
            evidence=[EvidenceRecord(kind="embedded_modules", data=data)],
        )
    except Exception as e:
        logger.error(f"MCP: Embed macros failed: {e}")
        return _operation_error(e)


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
        summary = validate_all_embedded_macros(WorkspacePathPolicy().input_file(ods_path))

        # Convert validation details to serializable format
        details = {}
        for module_name, result in summary.validation_details.items():
            details[module_name] = {
                "valid": result.valid,
                "errors": result.errors,
                "warnings": result.warnings,
                "functions_found": result.functions_found,
            }

        data = {
            "total_modules": summary.total_modules,
            "valid_syntax": summary.valid_syntax,
            "syntax_errors": summary.syntax_errors,
            "has_exported_scripts": summary.has_exported_scripts,
            "missing_exported_scripts": summary.missing_exported_scripts,
            "validation_details": details,
        }
        valid = summary.syntax_errors == 0 and summary.missing_exported_scripts == 0
        return _tool_response(
            GateExecutionStatus.PASSED if valid else GateExecutionStatus.FAILED,
            data=data,
            evidence=[EvidenceRecord(kind="macro_validation", data=data)],
            error=(
                None
                if valid
                else BoundaryError(
                    type="MacroValidationFailed",
                    message="One or more embedded macros failed validation",
                )
            ),
        )
    except Exception as e:
        logger.error(f"MCP: Validate macros failed: {e}")
        return _operation_error(e)


async def test_macro_execution(
    ods_path: str,
    script_uri: str,
    runtime_options: RuntimeToolOptions | None = None,
) -> dict[str, Any]:
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
        return _worker_tool_response(
            {"op": "execute_script", "ods_path": ods_path, "script_uri": script_uri},
            runtime_options,
        )
    except Exception as e:
        logger.error(f"MCP: Test macro execution failed: {e}")
        return _operation_error(e)


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
        script_infos = enumerate_python_scripts(WorkspacePathPolicy().input_file(ods_path))

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
                    "source_map_markers": info.source_map_markers or [],
                }
            )
            total_functions += len(info.functions)

        data = {
            "scripts": scripts,
            "total_scripts": len(scripts),
            "total_functions": total_functions,
        }
        return _tool_response(
            GateExecutionStatus.PASSED,
            data=data,
            evidence=[EvidenceRecord(kind="embedded_macro_inventory", data=data)],
        )
    except Exception as e:
        logger.error(f"MCP: List embedded macros failed: {e}")
        return _operation_error(e)


# ==============================================================================
# GUI Testing Operations
# ==============================================================================


async def open_document_gui(
    ods_path: str,
    use_xvfb: bool = True,
    keep_open: bool = True,
) -> dict[str, Any]:
    """Report GUI opening as unavailable across the Docker-only boundary.

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
    del use_xvfb, keep_open
    return _tool_response(
        GateExecutionStatus.UNAVAILABLE,
        implemented=False,
        capability_available=False,
        error=BoundaryError(
            type="CapabilityUnavailable",
            message=(
                f"Opening {Path(ods_path).name} on a host GUI is outside the "
                "Docker-only LibreOffice runtime"
            ),
        ),
    )


async def click_form_button(
    ods_path: str,
    button_name: str,
) -> dict[str, Any]:
    """Report real GUI/control clicking as unavailable.

    Args:
        ods_path: Path to ODS file
        button_name: Name of button control to click

    Returns:
        Dictionary with:
        - success: bool
        - message: str
        - error: str (if failed)
    """
    del ods_path, button_name
    return _tool_response(
        GateExecutionStatus.UNAVAILABLE,
        implemented=False,
        capability_available=False,
        error=BoundaryError(
            type="CapabilityUnavailable",
            message="Real GUI/control event dispatch is not implemented",
        ),
    )


async def execute_button_handler(
    ods_path: str,
    button_name: str,
    runtime_options: RuntimeToolOptions | None = None,
) -> dict[str, Any]:
    """Resolve a discovered button handler and invoke that script directly in Docker."""
    try:
        logger.info(f"MCP: Executing handler for button '{button_name}' in {ods_path}")
        return _worker_tool_response(
            {
                "op": "execute_button_handler",
                "ods_path": ods_path,
                "button_name": button_name,
                "use_gui": False,
            },
            runtime_options,
        )
    except Exception as exc:
        logger.error(f"MCP: Execute button handler failed: {exc}")
        return _operation_error(exc)


async def validate_document_runtime(
    ods_path: str, runtime_options: RuntimeToolOptions | None = None
) -> dict[str, Any]:
    """Run open/recalculate/save/close/reopen/package stages in the Docker target."""
    try:
        result = _worker_tool_response(
            {"op": "validate_document", "ods_path": ods_path}, runtime_options
        )
        if not result["success"]:
            return result
        stages = result.get("stages") or {}
        all_passed = bool(stages) and all(
            isinstance(stage, dict) and stage.get("status") == "passed" for stage in stages.values()
        )
        source_unchanged = not bool(result.get("source_mutated"))
        if all_passed and source_unchanged:
            return result
        result["success"] = False
        result["operation_status"] = GateExecutionStatus.FAILED.value
        result["error"] = BoundaryError(
            type="RuntimeValidationFailed",
            message="Not every required runtime stage passed or the source was mutated",
        ).model_dump(mode="json")
        return result
    except Exception as exc:
        return _operation_error(exc)


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
    logger.info(f"MCP: Keyboard input requested for {ods_path}")
    return _tool_response(
        GateExecutionStatus.UNAVAILABLE,
        implemented=False,
        capability_available=False,
        data={"keys_requested": len(key_sequence), "keys_sent": 0},
        error=BoundaryError(
            type="CapabilityUnavailable", message="Keyboard input is not implemented"
        ),
    )


async def get_cell_colors(
    ods_path: str,
    sheet_name: str,
    range_address: str,
    runtime_options: RuntimeToolOptions | None = None,
) -> dict[str, Any]:
    """Get background colors of cells in a range (useful for game state detection).

    Args:
        ods_path: Path to ODS file
        sheet_name: Sheet name
        range_address: Range address (for example, ``D3:M22``)

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

        return _worker_tool_response(
            {
                "op": "get_cell_colors",
                "ods_path": ods_path,
                "sheet_name": sheet_name,
                "range_address": range_address,
                "timeout_seconds": 30,
            },
            runtime_options,
        )

    except Exception as e:
        logger.error(f"MCP: Get cell colors failed: {e}")
        return _operation_error(e)


async def take_screenshot(
    ods_path: str,
    output_path: str,
    delay_seconds: float = 2.0,
) -> dict[str, Any]:
    """Report screenshot capture as unavailable until container capture exists.

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
    del ods_path, output_path, delay_seconds
    return _tool_response(
        GateExecutionStatus.UNAVAILABLE,
        implemented=False,
        capability_available=False,
        error=BoundaryError(
            type="CapabilityUnavailable",
            message="Screenshot capture is not implemented in the Docker runtime",
        ),
    )
