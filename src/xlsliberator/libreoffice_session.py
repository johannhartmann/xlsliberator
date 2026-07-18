"""Stateful, Docker-only LibreOffice session service."""

from __future__ import annotations

import hashlib
import json
import os
import shutil
import threading
from collections.abc import Callable
from dataclasses import dataclass, field
from datetime import UTC, datetime
from enum import StrEnum
from pathlib import Path
from typing import Any, Literal, Protocol
from uuid import uuid4

from pydantic import BaseModel, ConfigDict, Field

from xlsliberator.boundary_models import BoundaryError, BoundaryResponse, EvidenceRecord
from xlsliberator.docker_runtime import (
    DockerRuntimeIdentity,
    DockerRuntimeTimeout,
    DockerRuntimeUnavailable,
    LibreOfficeDockerRuntime,
)
from xlsliberator.execution_sandbox import WorkspacePathPolicy
from xlsliberator.scenarios.models import EnvironmentManifest, Scenario
from xlsliberator.validation_models import GateExecutionStatus


class StrictModel(BaseModel):
    """Versioned session boundary model."""

    model_config = ConfigDict(extra="forbid")


class SessionOperation(StrEnum):
    """Stable operation names exposed by the session service."""

    OPEN_DOCUMENT = "open_document"
    INSPECT_DOCUMENT = "inspect_document"
    LIST_SHEETS = "list_sheets"
    READ_CELLS = "read_cells"
    WRITE_CELLS = "write_cells"
    LIST_FORMULAS = "list_formulas"
    RECALCULATE = "recalculate"
    LIST_CONTROLS = "list_controls"
    DISPATCH_CONTROL_EVENT = "dispatch_control_event"
    SEND_KEYBOARD_EVENT = "send_keyboard_event"
    EXECUTE_PYTHON_MACRO = "execute_python_macro"
    CAPTURE_SCREENSHOT = "capture_screenshot"
    EXPORT_PDF = "export_pdf"
    SAVE = "save"
    CLOSE = "close"
    REOPEN = "reopen"
    RUN_SCENARIO = "run_scenario"


class SessionRuntimeDescriptor(StrictModel):
    """Exact runtime resources owned by one logical session."""

    schema_version: Literal["1.0.0"] = "1.0.0"
    session_id: str
    created_at: datetime
    state: str
    runtime_image: str
    runtime_image_digest: str
    libreoffice_version: str
    office_executable: str
    office_executable_sha256: str
    profile_identifier: str
    uno_port: int = Field(ge=1024, le=65535)
    display: str
    working_copy: str | None = None
    evidence_directory: str


class SessionBackendResult(StrictModel):
    """Truthful result from the Docker session backend."""

    transport_success: bool = True
    status: GateExecutionStatus
    implemented: bool = True
    capability_available: bool = True
    data: dict[str, Any] = Field(default_factory=dict)
    logs: list[str] = Field(default_factory=list)
    attachments: list[str] = Field(default_factory=list)
    error: BoundaryError | None = None


@dataclass(slots=True)
class SessionRecord:
    """Mutable state kept only by the trusted application-container service."""

    descriptor: SessionRuntimeDescriptor
    root: Path
    environment: EnvironmentManifest
    image_id: str
    working_copy: Path | None = None
    document_open: bool = False
    destroyed: bool = False
    failed: bool = False
    operation_index: int = 0
    lock: threading.RLock = field(default_factory=threading.RLock)


class SessionBackend(Protocol):
    """Runtime boundary implemented by Docker and deterministic fakes."""

    def create(
        self,
        *,
        session_id: str,
        profile_identifier: str,
        uno_port: int,
        display: str,
    ) -> SessionBackendResult:
        """Resolve and prove the exact runtime identity."""

    def execute(
        self,
        record: SessionRecord,
        operation: SessionOperation,
        parameters: dict[str, Any],
    ) -> SessionBackendResult:
        """Execute one operation against the session working copy."""

    def destroy(self, record: SessionRecord) -> SessionBackendResult:
        """Clean up every runtime process associated with the session."""


class DockerSessionBackend:
    """Execute session operations in isolated pinned LibreOffice containers."""

    _UI_UNAVAILABLE = {
        SessionOperation.DISPATCH_CONTROL_EVENT,
        SessionOperation.SEND_KEYBOARD_EVENT,
        SessionOperation.CAPTURE_SCREENSHOT,
    }

    def __init__(
        self,
        runtime: LibreOfficeDockerRuntime | None = None,
        *,
        timeout_seconds: float = 120,
    ) -> None:
        self.runtime = runtime or LibreOfficeDockerRuntime(timeout_seconds=timeout_seconds)
        self.timeout_seconds = timeout_seconds
        self._identity: DockerRuntimeIdentity | None = None

    def create(
        self,
        *,
        session_id: str,
        profile_identifier: str,
        uno_port: int,
        display: str,
    ) -> SessionBackendResult:
        del session_id, profile_identifier, uno_port, display
        try:
            identity = self.runtime.resolve_identity(probe=True)
        except DockerRuntimeUnavailable as exc:
            return _unavailable_backend_result(exc)
        self._identity = identity
        probe = identity.probe
        return SessionBackendResult(
            status=GateExecutionStatus.PASSED,
            data={
                "runtime_image": identity.image_reference,
                "runtime_image_digest": identity.image_id,
                "libreoffice_version": identity.version,
                "office_executable": str(probe.get("office_executable") or ""),
                "office_executable_sha256": str(probe.get("office_sha256") or ""),
                "architecture": identity.architecture,
                "python_version": probe.get("python_version"),
                "pyuno_identity": {
                    "uno_module": probe.get("uno_module"),
                    "uno_module_sha256": probe.get("uno_module_sha256"),
                    "pyuno_native_module": probe.get("pyuno_native_module"),
                    "pyuno_native_sha256": probe.get("pyuno_native_sha256"),
                },
            },
        )

    def execute(
        self,
        record: SessionRecord,
        operation: SessionOperation,
        parameters: dict[str, Any],
    ) -> SessionBackendResult:
        if operation in self._UI_UNAVAILABLE:
            return SessionBackendResult(
                status=GateExecutionStatus.UNAVAILABLE,
                implemented=False,
                capability_available=False,
                error=BoundaryError(
                    type="UIEventLayerUnavailable",
                    message=(
                        f"{operation.value} requires a verified real UI/event layer; "
                        "direct handler invocation is forbidden"
                    ),
                ),
            )
        if record.working_copy is None:
            return SessionBackendResult(
                status=GateExecutionStatus.NOT_RUN,
                error=BoundaryError(
                    type="DocumentNotOpen",
                    message="the session has no working document",
                ),
            )
        try:
            if operation is SessionOperation.LIST_SHEETS:
                return self._direct_request(record, "list_sheets", {})
            if operation is SessionOperation.READ_CELLS:
                return self._direct_request(
                    record,
                    "inspect_document_cells",
                    {"cells": list(parameters.get("cells") or [])},
                )
            if operation is SessionOperation.LIST_FORMULAS:
                return self._direct_request(record, "list_formula_cells", {})
            if operation is SessionOperation.INSPECT_DOCUMENT:
                return self._inspection(record)
            if operation is SessionOperation.LIST_CONTROLS:
                return self._controls(record)
            if operation is SessionOperation.OPEN_DOCUMENT:
                return self._open_or_reopen(record, "open")
            if operation is SessionOperation.CLOSE:
                return self._open_or_reopen(record, "close")
            if operation is SessionOperation.REOPEN:
                return self._open_or_reopen(record, "reopen")
            if operation is SessionOperation.WRITE_CELLS:
                return self._write_cells(record, parameters)
            if operation is SessionOperation.RECALCULATE:
                return self._mutating_actions(record, [{"kind": "recalculate"}])
            if operation is SessionOperation.EXECUTE_PYTHON_MACRO:
                return self._mutating_actions(
                    record,
                    [
                        {
                            "kind": "execute_python_macro",
                            "parameters": {"script_uri": str(parameters.get("script_uri") or "")},
                        }
                    ],
                )
            if operation is SessionOperation.SAVE:
                return self._mutating_actions(record, [{"kind": "save"}])
            if operation is SessionOperation.EXPORT_PDF:
                return self._export_pdf(record)
            if operation is SessionOperation.RUN_SCENARIO:
                scenario = Scenario.model_validate(parameters["scenario"])
                environment = EnvironmentManifest.model_validate(
                    parameters.get("environment") or record.environment
                )
                return self._scenario_request(
                    record,
                    scenario.model_dump(mode="json"),
                    environment,
                    mutate=True,
                )
        except (DockerRuntimeTimeout, DockerRuntimeUnavailable, OSError) as exc:
            return _unavailable_backend_result(exc)
        except Exception as exc:
            return SessionBackendResult(
                status=GateExecutionStatus.FAILED,
                error=BoundaryError(type=type(exc).__name__, message=str(exc)),
            )
        return SessionBackendResult(
            status=GateExecutionStatus.UNAVAILABLE,
            implemented=False,
            capability_available=False,
            error=BoundaryError(
                type="UnsupportedSessionOperation",
                message=f"session backend does not implement {operation.value}",
            ),
        )

    def destroy(self, record: SessionRecord) -> SessionBackendResult:
        del record
        return SessionBackendResult(
            status=GateExecutionStatus.PASSED,
            data={"process_tree_cleaned": True},
        )

    def _direct_request(
        self,
        record: SessionRecord,
        worker_op: str,
        parameters: dict[str, Any],
    ) -> SessionBackendResult:
        return self._request(
            record,
            {
                "op": worker_op,
                "ods_path": str(record.working_copy),
                **parameters,
            },
        )

    def _inspection(self, record: SessionRecord) -> SessionBackendResult:
        scenario = _scenario(
            "inspect-document",
            [
                {
                    "id": "open",
                    "action": {"kind": "open"},
                    "observations_after": [
                        {"id": "sheets", "kind": "sheet_state"},
                        {"id": "named_ranges", "kind": "named_ranges"},
                        {"id": "scripts", "kind": "embedded_scripts"},
                        {"id": "controls", "kind": "controls_bindings"},
                        {"id": "package_hash", "kind": "package_hash"},
                    ],
                }
            ],
        )
        return self._scenario_request(record, scenario, record.environment, mutate=False)

    def _controls(self, record: SessionRecord) -> SessionBackendResult:
        scenario = _scenario(
            "list-controls",
            [
                {
                    "id": "open",
                    "action": {"kind": "open"},
                    "observations_after": [
                        {"id": "controls", "kind": "controls_bindings"},
                    ],
                }
            ],
        )
        result = self._scenario_request(record, scenario, record.environment, mutate=False)
        if result.status is GateExecutionStatus.PASSED:
            result.data["controls"] = _observation_value(result.data, "open", "controls")
        return result

    def _open_or_reopen(self, record: SessionRecord, kind: str) -> SessionBackendResult:
        actions = [{"kind": "open"}]
        if kind == "close":
            actions.append({"kind": "close"})
        elif kind == "reopen":
            actions.extend([{"kind": "close"}, {"kind": "reopen"}])
        return self._scenario_request(
            record,
            _action_scenario(f"{kind}-document", actions),
            record.environment,
            mutate=False,
        )

    def _write_cells(
        self,
        record: SessionRecord,
        parameters: dict[str, Any],
    ) -> SessionBackendResult:
        actions: list[dict[str, Any]] = []
        for cell in parameters.get("cells") or []:
            if not isinstance(cell, dict):
                raise ValueError("write_cells entries must be objects")
            action_parameters = {
                key: cell[key]
                for key in ("sheet", "sheet_name", "address", "cell_address", "value", "formula")
                if key in cell
            }
            actions.append({"kind": "set_cell", "parameters": action_parameters})
        if not actions:
            raise ValueError("write_cells requires at least one cell update")
        actions.append({"kind": "save"})
        return self._mutating_actions(record, actions)

    def _mutating_actions(
        self,
        record: SessionRecord,
        actions: list[dict[str, Any]],
    ) -> SessionBackendResult:
        return self._scenario_request(
            record,
            _action_scenario("session-mutation", [{"kind": "open"}, *actions]),
            record.environment,
            mutate=True,
        )

    def _export_pdf(self, record: SessionRecord) -> SessionBackendResult:
        attachment = record.root / f"export-{record.operation_index:04d}.pdf"
        result = self._scenario_request(
            record,
            _action_scenario(
                "export-pdf",
                [{"kind": "open"}, {"kind": "export_pdf", "parameters": {"format": "pdf"}}],
            ),
            record.environment,
            mutate=False,
            attachment_output=attachment,
        )
        if result.status is GateExecutionStatus.PASSED:
            result.attachments.append(str(attachment))
            result.data["pdf"] = str(attachment)
            result.data["pdf_sha256"] = _hash_file(attachment)
        return result

    def _scenario_request(
        self,
        record: SessionRecord,
        scenario: dict[str, Any],
        environment: EnvironmentManifest,
        *,
        mutate: bool,
        attachment_output: Path | None = None,
    ) -> SessionBackendResult:
        working_copy = record.working_copy
        if working_copy is None:
            raise RuntimeError("session working copy is missing")
        candidate = record.root / f".working-{record.operation_index:04d}.ods"
        payload: dict[str, Any] = {
            "op": "run_scenario",
            "ods_path": str(working_copy),
            "scenario": scenario,
            "environment": environment.model_dump(mode="json"),
            "final_save_reopen": mutate,
        }
        if mutate:
            payload["output_path"] = str(candidate)
        if attachment_output is not None:
            payload["attachment_output_path"] = str(attachment_output)
        result = self._request(record, payload)
        if mutate:
            if result.status is GateExecutionStatus.PASSED and candidate.is_file():
                os.replace(candidate, working_copy)
                result.data["working_copy_sha256"] = _hash_file(working_copy)
            else:
                candidate.unlink(missing_ok=True)
        return result

    def _request(
        self,
        record: SessionRecord,
        payload: dict[str, Any],
    ) -> SessionBackendResult:
        identity = self._identity
        if identity is None or identity.image_id != record.image_id:
            raise DockerRuntimeUnavailable("session runtime identity is unavailable or changed")
        request = {
            **payload,
            "session_profile_identifier": record.descriptor.profile_identifier,
            "session_port": record.descriptor.uno_port,
            "session_display": record.descriptor.display,
            "timeout_seconds": self.timeout_seconds,
        }
        response = self.runtime.request(request, _identity=record.image_id)
        data = dict(response.get("data") or {})
        logs = [str(item) for item in data.get("logs") or []]
        stderr = str(data.get("container_stderr") or "")
        if stderr:
            logs.append(stderr)
        if not response.get("success"):
            raw_error = dict(response.get("error") or {})
            return SessionBackendResult(
                transport_success=False,
                status=GateExecutionStatus.FAILED,
                data=data,
                logs=logs,
                error=BoundaryError(
                    type=str(raw_error.get("type") or "WorkerError"),
                    message=str(raw_error.get("message") or "LibreOffice worker failed"),
                    details=raw_error,
                ),
            )
        raw_status = str(data.get("status") or GateExecutionStatus.PASSED.value)
        try:
            status = GateExecutionStatus(raw_status)
        except ValueError:
            status = GateExecutionStatus.FAILED
            return SessionBackendResult(
                status=status,
                data=data,
                logs=logs,
                error=BoundaryError(
                    type="MalformedOperationStatus",
                    message=f"LibreOffice worker returned unknown status {raw_status!r}",
                ),
            )
        error = None
        if status is not GateExecutionStatus.PASSED:
            error = BoundaryError(
                type="LibreOfficeOperationFailed",
                message=f"LibreOffice operation was {status.value}",
            )
        return SessionBackendResult(status=status, data=data, logs=logs, error=error)


class FakeSessionBackend:
    """Deterministic backend for server/client and state-machine tests."""

    def __init__(
        self,
        *,
        statuses: dict[SessionOperation, GateExecutionStatus] | None = None,
    ) -> None:
        self.statuses = dict(statuses or {})
        self.calls: list[tuple[str, str, dict[str, Any]]] = []
        self.cells: dict[str, dict[tuple[str, str], Any]] = {}

    def create(
        self,
        *,
        session_id: str,
        profile_identifier: str,
        uno_port: int,
        display: str,
    ) -> SessionBackendResult:
        self.calls.append(
            (
                session_id,
                "create_session",
                {
                    "profile_identifier": profile_identifier,
                    "uno_port": uno_port,
                    "display": display,
                },
            )
        )
        self.cells[session_id] = {("Sheet1", "A1"): 1.0}
        return SessionBackendResult(
            status=GateExecutionStatus.PASSED,
            data={
                "runtime_image": "fake-libreoffice:26.2.4.2",
                "runtime_image_digest": "sha256:" + "f" * 64,
                "libreoffice_version": "26.2.4.2",
                "office_executable": "/opt/libreoffice26.2/program/soffice",
                "office_executable_sha256": "e" * 64,
            },
        )

    def execute(
        self,
        record: SessionRecord,
        operation: SessionOperation,
        parameters: dict[str, Any],
    ) -> SessionBackendResult:
        self.calls.append((record.descriptor.session_id, operation.value, dict(parameters)))
        if operation in DockerSessionBackend._UI_UNAVAILABLE:
            return SessionBackendResult(
                status=GateExecutionStatus.UNAVAILABLE,
                implemented=False,
                capability_available=False,
                error=BoundaryError(
                    type="UIEventLayerUnavailable",
                    message=f"fake backend has no verified {operation.value} UI layer",
                ),
            )
        status = self.statuses.get(operation, GateExecutionStatus.PASSED)
        if status is not GateExecutionStatus.PASSED:
            return SessionBackendResult(
                status=status,
                capability_available=status is not GateExecutionStatus.UNAVAILABLE,
                error=BoundaryError(type="FakeFailure", message=f"fake {status.value}"),
            )
        data: dict[str, Any] = {"operation": operation.value}
        cells = self.cells[record.descriptor.session_id]
        if operation is SessionOperation.LIST_SHEETS:
            data.update({"sheets": ["Sheet1"], "count": 1})
        elif operation is SessionOperation.WRITE_CELLS:
            for item in parameters.get("cells") or []:
                sheet = str(item.get("sheet") or item.get("sheet_name") or "Sheet1")
                address = str(item.get("address") or item.get("cell_address") or "")
                cells[(sheet, address)] = item.get("formula", item.get("value"))
            data["written"] = len(parameters.get("cells") or [])
        elif operation is SessionOperation.READ_CELLS:
            data["cells"] = [
                {
                    **item,
                    "value": cells.get(
                        (
                            str(item.get("sheet") or "Sheet1"),
                            str(item.get("address") or ""),
                        )
                    ),
                }
                for item in parameters.get("cells") or []
            ]
        elif operation is SessionOperation.LIST_FORMULAS:
            formulas = [
                {"sheet": sheet, "address": address, "formula": value}
                for (sheet, address), value in sorted(cells.items())
                if isinstance(value, str) and value.startswith("=")
            ]
            data.update({"cells": formulas, "count": len(formulas)})
        elif operation is SessionOperation.EXPORT_PDF:
            attachment = record.root / f"fake-export-{record.operation_index:04d}.pdf"
            attachment.write_bytes(b"%PDF-1.4\n% deterministic fake\n")
            data["pdf"] = str(attachment)
            return SessionBackendResult(
                status=status,
                data=data,
                attachments=[str(attachment)],
                logs=[f"{operation.value}:passed"],
            )
        elif operation is SessionOperation.RUN_SCENARIO:
            scenario = dict(parameters["scenario"])
            timestamp = "2000-01-01T00:00:00+00:00"
            data.update(
                {
                    "scenario_id": str(scenario.get("id") or ""),
                    "status": "passed",
                    "started_at": timestamp,
                    "ended_at": timestamp,
                    "steps": [
                        {
                            "step_id": str(step["id"]),
                            "action": str(dict(step["action"])["kind"]),
                            "status": "passed",
                            "started_at": timestamp,
                            "ended_at": timestamp,
                        }
                        for step in scenario.get("steps") or []
                    ],
                    "runtime": {
                        "profile_identifier": record.descriptor.profile_identifier,
                        "uno_port": record.descriptor.uno_port,
                        "display": record.descriptor.display,
                    },
                }
            )
        return SessionBackendResult(
            status=status,
            data=data,
            logs=[f"{operation.value}:{status.value}"],
        )

    def destroy(self, record: SessionRecord) -> SessionBackendResult:
        self.calls.append((record.descriptor.session_id, "destroy_session", {}))
        return SessionBackendResult(
            status=GateExecutionStatus.PASSED,
            data={"process_tree_cleaned": True},
        )


class LibreOfficeSessionManager:
    """Own state, working copies, logs, and lifecycle for LibreOffice sessions."""

    def __init__(
        self,
        *,
        backend: SessionBackend | None = None,
        workspace: WorkspacePathPolicy | None = None,
        session_root: Path | None = None,
        id_factory: Callable[[], str] | None = None,
        clock: Callable[[], datetime] | None = None,
        first_port: int = 22000,
    ) -> None:
        self.backend = backend or DockerSessionBackend()
        self.workspace = workspace or WorkspacePathPolicy()
        configured = os.environ.get("XLSLIBERATOR_SESSION_ROOT")
        root = session_root or (
            Path(configured) if configured else Path.cwd() / "artifacts/sessions"
        )
        root_parent = root.expanduser().resolve().parent
        if not root_parent.exists():
            root_parent.mkdir(parents=True, exist_ok=True)
        self.session_root = root.expanduser().resolve()
        if not any(
            self.session_root == allowed or self.session_root.is_relative_to(allowed)
            for allowed in self.workspace.roots
        ):
            raise ValueError("session root must be beneath a configured workspace root")
        self.session_root.mkdir(parents=True, exist_ok=True)
        self._id_factory = id_factory or (lambda: uuid4().hex)
        self._clock = clock or (lambda: datetime.now(UTC))
        self._next_port = first_port
        self._sessions: dict[str, SessionRecord] = {}
        self._archived: dict[str, SessionRecord] = {}
        self._lock = threading.RLock()

    def create_session(
        self,
        *,
        environment: EnvironmentManifest | dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        typed_environment = EnvironmentManifest.model_validate(environment or {})
        session_id = self._id_factory()
        if not session_id or any(
            character not in "abcdefghijklmnopqrstuvwxyz0123456789-" for character in session_id
        ):
            raise ValueError("session ID factory returned a malformed identifier")
        with self._lock:
            if session_id in self._sessions or session_id in self._archived:
                raise ValueError(f"session ID already exists: {session_id}")
            port = self._allocate_port()
            root = self.session_root / session_id
            root.mkdir(mode=0o700)
        profile = f"xlsliberator-session-{session_id}"
        display = f":{port - 20000}"
        try:
            result = self.backend.create(
                session_id=session_id,
                profile_identifier=profile,
                uno_port=port,
                display=display,
            )
        except Exception as exc:
            result = _unavailable_backend_result(exc)
        descriptor = SessionRuntimeDescriptor(
            session_id=session_id,
            created_at=self._clock(),
            state=("created" if result.status is GateExecutionStatus.PASSED else "creation_failed"),
            runtime_image=str(result.data.get("runtime_image") or "unavailable"),
            runtime_image_digest=str(result.data.get("runtime_image_digest") or "unavailable"),
            libreoffice_version=str(result.data.get("libreoffice_version") or "unavailable"),
            office_executable=str(result.data.get("office_executable") or "unavailable"),
            office_executable_sha256=str(
                result.data.get("office_executable_sha256") or "unavailable"
            ),
            profile_identifier=profile,
            uno_port=port,
            display=display,
            evidence_directory=str(root),
        )
        record = SessionRecord(
            descriptor=descriptor,
            root=root,
            environment=typed_environment,
            image_id=descriptor.runtime_image_digest,
            failed=result.status is not GateExecutionStatus.PASSED,
        )
        with self._lock:
            if result.status is GateExecutionStatus.PASSED:
                self._sessions[session_id] = record
            else:
                record.destroyed = True
                self._archived[session_id] = record
        self._record(record, "create_session", {}, result)
        return _boundary_payload(
            result,
            session_id=session_id,
            descriptor=descriptor.model_dump(mode="json"),
        )

    def open_document(self, session_id: str, document_path: str | Path) -> dict[str, Any]:
        record = self._active(session_id)
        source = self.workspace.input_file(document_path)
        with record.lock:
            if record.document_open:
                return self._state_failure(record, "document is already open")
            working_copy = record.root / "working-copy.ods"
            _atomic_copy(source, working_copy)
            record.working_copy = working_copy
            record.descriptor.working_copy = str(working_copy)
            result = self._execute(record, SessionOperation.OPEN_DOCUMENT, {})
            if result.status is GateExecutionStatus.PASSED:
                record.document_open = True
                record.descriptor.state = "open"
            return self._finish(record, SessionOperation.OPEN_DOCUMENT, {}, result)

    def perform(
        self,
        session_id: str,
        operation: SessionOperation,
        **parameters: Any,
    ) -> dict[str, Any]:
        if operation is SessionOperation.OPEN_DOCUMENT:
            raise ValueError("use open_document with an explicit workspace path")
        record = self._active(session_id)
        with record.lock:
            if operation is SessionOperation.REOPEN:
                if record.document_open:
                    return self._state_failure(record, "document must be closed before reopen")
            elif not record.document_open:
                return self._state_failure(record, "document is not open")
            result = self._execute(record, operation, dict(parameters))
            if result.status is GateExecutionStatus.PASSED:
                if operation is SessionOperation.CLOSE:
                    record.document_open = False
                    record.descriptor.state = "closed"
                elif operation is SessionOperation.REOPEN:
                    record.document_open = True
                    record.descriptor.state = "open"
                if operation is SessionOperation.SAVE and parameters.get("output_path"):
                    destination = self.workspace.output_file(str(parameters["output_path"]))
                    if record.working_copy is None:
                        return self._state_failure(record, "session working copy is missing")
                    _atomic_copy(record.working_copy, destination)
                    result.data["saved_path"] = str(destination)
                    result.data["saved_sha256"] = _hash_file(destination)
                if operation is SessionOperation.EXPORT_PDF and parameters.get("output_path"):
                    destination = self.workspace.output_file(str(parameters["output_path"]))
                    if len(result.attachments) != 1:
                        return self._state_failure(
                            record,
                            "LibreOffice did not return exactly one PDF attachment",
                        )
                    _atomic_copy(Path(result.attachments[0]), destination)
                    result.data["pdf"] = str(destination)
                    result.data["pdf_sha256"] = _hash_file(destination)
            return self._finish(record, operation, dict(parameters), result)

    def run_scenario(
        self,
        session_id: str,
        *,
        scenario: Scenario,
        environment: EnvironmentManifest,
    ) -> dict[str, Any]:
        return self.perform(
            session_id,
            SessionOperation.RUN_SCENARIO,
            scenario=scenario.model_dump(mode="json"),
            environment=environment.model_dump(mode="json"),
        )

    def collect_logs(self, session_id: str) -> dict[str, Any]:
        record = self._lookup(session_id)
        with record.lock:
            operation_log = record.root / "operations.jsonl"
            office_log = record.root / "office.log"
            attachments = sorted(
                str(path)
                for path in record.root.iterdir()
                if path.is_file()
                and path.name not in {"operations.jsonl", "office.log", "working-copy.ods"}
                and not path.name.startswith(".")
            )
            result = SessionBackendResult(
                status=GateExecutionStatus.PASSED,
                data={
                    "state": record.descriptor.state,
                    "failed": record.failed,
                    "operation_log": (
                        operation_log.read_text(encoding="utf-8") if operation_log.is_file() else ""
                    ),
                    "office_log": (
                        office_log.read_text(encoding="utf-8") if office_log.is_file() else ""
                    ),
                    "attachments": attachments,
                    "evidence_directory": str(record.root),
                },
            )
            return _boundary_payload(result, session_id=session_id)

    def destroy_session(self, session_id: str) -> dict[str, Any]:
        record = self._active(session_id)
        with record.lock:
            result = self.backend.destroy(record)
            if result.status is GateExecutionStatus.PASSED:
                record.destroyed = True
                record.document_open = False
                record.descriptor.state = "destroyed"
                with self._lock:
                    self._sessions.pop(session_id, None)
                    self._archived[session_id] = record
            self._record(record, "destroy_session", {}, result)
            result.data.setdefault("logs_preserved", True)
            result.data.setdefault("evidence_directory", str(record.root))
            return _boundary_payload(result, session_id=session_id)

    def descriptor(self, session_id: str) -> SessionRuntimeDescriptor:
        return self._lookup(session_id).descriptor.model_copy(deep=True)

    def _execute(
        self,
        record: SessionRecord,
        operation: SessionOperation,
        parameters: dict[str, Any],
    ) -> SessionBackendResult:
        record.operation_index += 1
        return self.backend.execute(record, operation, parameters)

    def _finish(
        self,
        record: SessionRecord,
        operation: SessionOperation,
        parameters: dict[str, Any],
        result: SessionBackendResult,
    ) -> dict[str, Any]:
        if result.status is not GateExecutionStatus.PASSED:
            record.failed = True
        self._record(record, operation.value, parameters, result)
        return _boundary_payload(
            result,
            session_id=record.descriptor.session_id,
            session_state=record.descriptor.state,
            descriptor=record.descriptor.model_dump(mode="json"),
        )

    def _record(
        self,
        record: SessionRecord,
        operation: str,
        parameters: dict[str, Any],
        result: SessionBackendResult,
    ) -> None:
        entry = {
            "index": record.operation_index,
            "operation": operation,
            "parameters": _redacted_parameters(parameters),
            "transport_success": result.transport_success,
            "operation_status": result.status.value,
            "implemented": result.implemented,
            "capability_available": result.capability_available,
            "error": result.error.model_dump(mode="json") if result.error else None,
        }
        with (record.root / "operations.jsonl").open("a", encoding="utf-8") as handle:
            handle.write(json.dumps(entry, sort_keys=True, separators=(",", ":")) + "\n")
            handle.flush()
            os.fsync(handle.fileno())
        _append_backend_logs(record.root, result.logs)

    def _state_failure(self, record: SessionRecord, message: str) -> dict[str, Any]:
        result = SessionBackendResult(
            status=GateExecutionStatus.NOT_RUN,
            error=BoundaryError(type="InvalidSessionState", message=message),
        )
        record.failed = True
        self._record(record, "state_validation", {}, result)
        return _boundary_payload(
            result,
            session_id=record.descriptor.session_id,
            session_state=record.descriptor.state,
        )

    def _active(self, session_id: str) -> SessionRecord:
        with self._lock:
            record = self._sessions.get(session_id)
        if record is None:
            raise KeyError(f"active LibreOffice session not found: {session_id}")
        return record

    def _lookup(self, session_id: str) -> SessionRecord:
        with self._lock:
            record = self._sessions.get(session_id) or self._archived.get(session_id)
        if record is None:
            raise KeyError(f"LibreOffice session not found: {session_id}")
        return record

    def _allocate_port(self) -> int:
        used = {record.descriptor.uno_port for record in self._sessions.values()}
        for _ in range(40000):
            candidate = self._next_port
            self._next_port += 1
            if self._next_port > 62000:
                self._next_port = 22000
            if candidate not in used:
                return candidate
        raise RuntimeError("no LibreOffice session port is available")


def _scenario(identifier: str, steps: list[dict[str, Any]]) -> dict[str, Any]:
    return {"schema_version": "1.0.0", "id": identifier, "steps": steps}


def _action_scenario(identifier: str, actions: list[dict[str, Any]]) -> dict[str, Any]:
    return _scenario(
        identifier,
        [
            {
                "id": f"{index:02d}-{str(action['kind'])}",
                "action": {
                    "kind": str(action["kind"]),
                    "parameters": dict(action.get("parameters") or {}),
                    "required": True,
                },
            }
            for index, action in enumerate(actions, start=1)
        ],
    )


def _observation_value(
    data: dict[str, Any],
    step_id: str,
    observation_id: str,
) -> Any:
    for step in data.get("steps") or []:
        if isinstance(step, dict) and step.get("step_id") == step_id:
            observation = dict(step.get("observations_after") or {}).get(observation_id)
            if isinstance(observation, dict):
                return observation.get("value")
    return None


def _unavailable_backend_result(exc: BaseException) -> SessionBackendResult:
    return SessionBackendResult(
        transport_success=False,
        status=GateExecutionStatus.UNAVAILABLE,
        capability_available=False,
        error=BoundaryError(type=type(exc).__name__, message=str(exc)),
    )


def _boundary_payload(
    result: SessionBackendResult,
    *,
    session_id: str,
    **data: Any,
) -> dict[str, Any]:
    merged = {"session_id": session_id, **result.data, **data}
    return BoundaryResponse(
        transport_success=result.transport_success,
        operation_status=result.status,
        implemented=result.implemented,
        capability_available=result.capability_available,
        evidence=[
            *(
                [EvidenceRecord(kind="session_logs", data={"items": result.logs})]
                if result.logs
                else []
            ),
            *[EvidenceRecord(kind="attachment", path=path) for path in result.attachments],
        ],
        error=result.error,
        data=merged,
    ).to_payload()


def _append_backend_logs(root: Path, logs: list[str]) -> None:
    if not logs:
        return
    with (root / "office.log").open("a", encoding="utf-8") as handle:
        for item in logs:
            handle.write(item.rstrip() + "\n")
        handle.flush()
        os.fsync(handle.fileno())


def _atomic_copy(source: Path, destination: Path) -> None:
    temporary = destination.with_name(f".{destination.name}.{uuid4().hex}")
    try:
        shutil.copy2(source, temporary)
        os.replace(temporary, destination)
    finally:
        temporary.unlink(missing_ok=True)


def _hash_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _redacted_parameters(parameters: dict[str, Any]) -> dict[str, Any]:
    redacted = dict(parameters)
    if "environment" in redacted:
        redacted["environment"] = "<typed-environment>"
    if "scenario" in redacted:
        scenario = redacted["scenario"]
        redacted["scenario"] = (
            {"id": scenario.get("id")} if isinstance(scenario, dict) else "<typed-scenario>"
        )
    return redacted
