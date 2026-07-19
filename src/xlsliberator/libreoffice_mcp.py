"""Session-oriented MCP tools for the LibreOffice-only runtime."""

from __future__ import annotations

import threading
from pathlib import Path
from typing import Any

from xlsliberator.application_showcase import (
    build_candidate,
    bundle_application_replays as bundle_replays,
    run_application_scenario as run_scenario,
)
from xlsliberator.boundary_models import BoundaryError, BoundaryResponse
from xlsliberator.container_boundary import require_application_container
from xlsliberator.execution_sandbox import WorkspacePathPolicy
from xlsliberator.libreoffice_session import LibreOfficeSessionManager, SessionOperation
from xlsliberator.validation_models import GateExecutionStatus

_manager: LibreOfficeSessionManager | None = None
_manager_lock = threading.Lock()


def configure_session_manager(manager: LibreOfficeSessionManager | None) -> None:
    """Replace the process-local manager for deterministic tests."""
    global _manager
    with _manager_lock:
        _manager = manager


def get_session_manager() -> LibreOfficeSessionManager:
    """Return the trusted application-container session manager."""
    require_application_container()
    global _manager
    with _manager_lock:
        if _manager is None:
            _manager = LibreOfficeSessionManager()
        return _manager


async def create_session(environment: dict[str, Any] | None = None) -> dict[str, Any]:
    """Create an isolated LibreOffice session and return its explicit session ID."""
    return _call(lambda: get_session_manager().create_session(environment=environment))


async def open_document(session_id: str, document_path: str) -> dict[str, Any]:
    """Copy an ODS document into the session and open the disposable copy."""
    return _call(
        lambda: get_session_manager().open_document(session_id, document_path),
        session_id=session_id,
    )


async def inspect_document(session_id: str) -> dict[str, Any]:
    """Inspect the currently open session document."""
    return _perform(session_id, SessionOperation.INSPECT_DOCUMENT)


async def list_sheets(session_id: str) -> dict[str, Any]:
    """List sheets in the currently open session document."""
    return _perform(session_id, SessionOperation.LIST_SHEETS)


async def read_cells(session_id: str, cells: list[dict[str, Any]]) -> dict[str, Any]:
    """Read typed cell values, formulas, types, and errors."""
    return _perform(session_id, SessionOperation.READ_CELLS, cells=cells)


async def write_cells(session_id: str, cells: list[dict[str, Any]]) -> dict[str, Any]:
    """Write explicit cell values or formulas to the session working copy."""
    return _perform(session_id, SessionOperation.WRITE_CELLS, cells=cells)


async def list_formulas(session_id: str) -> dict[str, Any]:
    """List formulas and evaluated states in the current document."""
    return _perform(session_id, SessionOperation.LIST_FORMULAS)


async def recalculate(session_id: str) -> dict[str, Any]:
    """Recalculate and persist the session working copy."""
    return _perform(session_id, SessionOperation.RECALCULATE)


async def list_controls(session_id: str) -> dict[str, Any]:
    """List ODF controls and their event bindings."""
    return _perform(session_id, SessionOperation.LIST_CONTROLS)


async def dispatch_control_event(session_id: str, control_name: str) -> dict[str, Any]:
    """Dispatch a real UI control event or truthfully return UNAVAILABLE."""
    return _perform(
        session_id,
        SessionOperation.DISPATCH_CONTROL_EVENT,
        control_name=control_name,
    )


async def send_keyboard_event(session_id: str, keys: list[str]) -> dict[str, Any]:
    """Dispatch keyboard input with proof, or truthfully return UNAVAILABLE."""
    return _perform(session_id, SessionOperation.SEND_KEYBOARD_EVENT, keys=keys)


async def execute_python_macro(session_id: str, script_uri: str) -> dict[str, Any]:
    """Execute one embedded Python macro through LibreOffice's scripting provider."""
    return _perform(
        session_id,
        SessionOperation.EXECUTE_PYTHON_MACRO,
        script_uri=script_uri,
    )


async def capture_screenshot(session_id: str) -> dict[str, Any]:
    """Capture a real UI screenshot or truthfully return UNAVAILABLE."""
    return _perform(session_id, SessionOperation.CAPTURE_SCREENSHOT)


async def export_pdf(session_id: str, output_path: str) -> dict[str, Any]:
    """Export the current session document to a workspace-confined PDF."""
    return _perform(session_id, SessionOperation.EXPORT_PDF, output_path=output_path)


async def save(session_id: str, output_path: str | None = None) -> dict[str, Any]:
    """Persist the session working copy and optionally copy it to a workspace path."""
    return _perform(session_id, SessionOperation.SAVE, output_path=output_path)


async def close(session_id: str) -> dict[str, Any]:
    """Close the logical session document."""
    return _perform(session_id, SessionOperation.CLOSE)


async def reopen(session_id: str) -> dict[str, Any]:
    """Reopen the closed session working copy through LibreOffice."""
    return _perform(session_id, SessionOperation.REOPEN)


async def collect_logs(session_id: str) -> dict[str, Any]:
    """Collect preserved logs and attachments, including after destruction."""
    return _call(
        lambda: get_session_manager().collect_logs(session_id),
        session_id=session_id,
    )


async def destroy_session(session_id: str) -> dict[str, Any]:
    """Clean the runtime process tree while preserving session evidence."""
    return _call(
        lambda: get_session_manager().destroy_session(session_id),
        session_id=session_id,
    )


async def build_application_candidate(
    source_path: str,
    candidate_path: str,
    output_path: str,
) -> dict[str, Any]:
    """Build an ODS through a content-bound generated migration candidate."""

    def operation() -> dict[str, Any]:
        workspace = WorkspacePathPolicy()
        source = workspace.input_file(source_path)
        candidate = workspace.input_file(candidate_path)
        output = _migration_output(workspace, output_path, ".ods")
        return _migration_payload(build_candidate(source, candidate, output))

    return _call(operation)


async def run_application_scenario(
    target_path: str,
    candidate_path: str,
    evidence_path: str,
    scenario_id: str,
    actions: list[dict[str, Any]],
    adapter_config: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Run one declared real-GUI scenario and retain its replay evidence."""

    def operation() -> dict[str, Any]:
        workspace = WorkspacePathPolicy()
        target = workspace.input_file(target_path)
        candidate = workspace.input_file(candidate_path)
        evidence = _migration_output(workspace, evidence_path, ".zip")
        return _migration_payload(
            run_scenario(
                target,
                candidate,
                evidence,
                actions,
                scenario_id=scenario_id,
                adapter_config=adapter_config,
            )
        )

    return _call(operation)


async def bundle_application_replays(
    evidence_paths: dict[str, str],
    output_path: str,
    replay_id: str,
) -> dict[str, Any]:
    """Bundle a declared scenario set into a recorded, browser-replayable demo."""

    def operation() -> dict[str, Any]:
        workspace = WorkspacePathPolicy()
        evidence = {
            scenario_id: workspace.input_file(path)
            for scenario_id, path in evidence_paths.items()
        }
        output = _migration_output(workspace, output_path, ".zip")
        return _migration_payload(bundle_replays(evidence, output, replay_id=replay_id))

    return _call(operation)


def _migration_output(
    workspace: WorkspacePathPolicy,
    output_path: str,
    suffix: str,
) -> Path:
    output = workspace.output_file(output_path)
    if output.suffix.lower() != suffix:
        raise ValueError(f"migration output must use the {suffix} suffix")
    return output


def _migration_payload(result: dict[str, Any]) -> dict[str, Any]:
    return BoundaryResponse(
        transport_success=True,
        operation_status=GateExecutionStatus.PASSED,
        implemented=True,
        capability_available=True,
        data=result,
    ).to_payload()


def _perform(
    session_id: str,
    operation: SessionOperation,
    **parameters: Any,
) -> dict[str, Any]:
    return _call(
        lambda: get_session_manager().perform(
            session_id,
            operation,
            **parameters,
        ),
        session_id=session_id,
    )


def _call(callback: Any, *, session_id: str | None = None) -> dict[str, Any]:
    try:
        result = callback()
        if not isinstance(result, dict):
            raise TypeError("session manager returned a non-object response")
        return result
    except Exception as exc:
        response = BoundaryResponse(
            transport_success=False,
            operation_status=GateExecutionStatus.FAILED,
            implemented=True,
            capability_available=True,
            error=BoundaryError(type=type(exc).__name__, message=str(exc)),
            data={"session_id": session_id} if session_id is not None else {},
        )
        return response.to_payload()
