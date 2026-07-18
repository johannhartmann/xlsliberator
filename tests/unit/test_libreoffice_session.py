"""Stateful LibreOffice MCP session and deterministic client tests."""

from __future__ import annotations

import asyncio
import hashlib
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

import pytest

from xlsliberator.cli import cli
from xlsliberator.execution_sandbox import WorkspacePathPolicy
from xlsliberator.libreoffice_mcp import (
    close,
    collect_logs,
    configure_session_manager,
    create_session,
    destroy_session,
    list_sheets,
    open_document,
)
from xlsliberator.libreoffice_session import (
    FakeSessionBackend,
    LibreOfficeSessionManager,
    SessionBackendResult,
    SessionOperation,
)
from xlsliberator.libreoffice_session_client import (
    InProcessSessionTransport,
    LibreOfficeSessionClient,
)
from xlsliberator.libreoffice_session_scenario_runner import (
    LibreOfficeSessionScenarioRunner,
)
from xlsliberator.scenarios.models import (
    Action,
    ActionKind,
    EnvironmentManifest,
    Scenario,
    ScenarioStep,
)
from xlsliberator.validation_models import GateExecutionStatus


@pytest.fixture
def manager(tmp_path: Path) -> LibreOfficeSessionManager:
    return LibreOfficeSessionManager(
        backend=FakeSessionBackend(),
        workspace=WorkspacePathPolicy([tmp_path]),
        session_root=tmp_path / "sessions",
        id_factory=lambda: "session-001",
        clock=lambda: datetime(2000, 1, 1, tzinfo=UTC),
    )


def test_session_lifecycle_owns_explicit_runtime_resources_and_working_copy(
    manager: LibreOfficeSessionManager,
    tmp_path: Path,
) -> None:
    source = tmp_path / "source.ods"
    source.write_bytes(b"original")
    original_hash = hashlib.sha256(source.read_bytes()).hexdigest()

    created = manager.create_session()
    session_id = created["session_id"]
    opened = manager.open_document(session_id, source)
    sheets = manager.perform(session_id, SessionOperation.LIST_SHEETS)
    written = manager.perform(
        session_id,
        SessionOperation.WRITE_CELLS,
        cells=[{"sheet": "Sheet1", "address": "B2", "formula": "=1+1"}],
    )
    formulas = manager.perform(session_id, SessionOperation.LIST_FORMULAS)
    saved = manager.perform(
        session_id,
        SessionOperation.SAVE,
        output_path=str(tmp_path / "saved.ods"),
    )
    closed = manager.perform(session_id, SessionOperation.CLOSE)
    reopened = manager.perform(session_id, SessionOperation.REOPEN)
    exported = manager.perform(
        session_id,
        SessionOperation.EXPORT_PDF,
        output_path=str(tmp_path / "export.pdf"),
    )

    descriptor = created["descriptor"]
    assert created["transport_success"] is True
    assert created["operation_status"] == "passed"
    assert descriptor["profile_identifier"] == "xlsliberator-session-session-001"
    assert descriptor["uno_port"] == 22000
    assert descriptor["display"] == ":2000"
    assert descriptor["libreoffice_version"] == "26.2.4.2"
    assert opened["session_id"] == session_id
    assert opened["session_state"] == "open"
    assert sheets["sheets"] == ["Sheet1"]
    assert written["success"] is True
    assert formulas["cells"] == [{"sheet": "Sheet1", "address": "B2", "formula": "=1+1"}]
    assert saved["saved_path"] == str(tmp_path / "saved.ods")
    assert closed["session_state"] == "closed"
    assert reopened["session_state"] == "open"
    assert Path(exported["pdf"]).read_bytes().startswith(b"%PDF")
    assert hashlib.sha256(source.read_bytes()).hexdigest() == original_hash


def test_ui_and_keyboard_operations_are_truthfully_unavailable(
    manager: LibreOfficeSessionManager,
    tmp_path: Path,
) -> None:
    source = tmp_path / "source.ods"
    source.write_bytes(b"original")
    session_id = manager.create_session()["session_id"]
    manager.open_document(session_id, source)

    control = manager.perform(
        session_id,
        SessionOperation.DISPATCH_CONTROL_EVENT,
        control_name="Run",
    )
    keyboard = manager.perform(
        session_id,
        SessionOperation.SEND_KEYBOARD_EVENT,
        keys=["ENTER"],
    )
    screenshot = manager.perform(session_id, SessionOperation.CAPTURE_SCREENSHOT)

    for result in (control, keyboard, screenshot):
        assert result["success"] is False
        assert result["implemented"] is False
        assert result["capability_available"] is False
        assert result["operation_status"] == "unavailable"
        assert result["error"]["type"] == "UIEventLayerUnavailable"


def test_sessions_have_isolated_resources_and_working_state(tmp_path: Path) -> None:
    identifiers = iter(("session-a", "session-b"))
    manager = LibreOfficeSessionManager(
        backend=FakeSessionBackend(),
        workspace=WorkspacePathPolicy([tmp_path]),
        session_root=tmp_path / "sessions",
        id_factory=lambda: next(identifiers),
    )
    source = tmp_path / "source.ods"
    source.write_bytes(b"source")

    first = manager.create_session()
    second = manager.create_session()
    manager.open_document(first["session_id"], source)
    manager.open_document(second["session_id"], source)
    manager.perform(
        first["session_id"],
        SessionOperation.WRITE_CELLS,
        cells=[{"sheet": "Sheet1", "address": "C3", "value": 99}],
    )
    first_value = manager.perform(
        first["session_id"],
        SessionOperation.READ_CELLS,
        cells=[{"sheet": "Sheet1", "address": "C3"}],
    )
    second_value = manager.perform(
        second["session_id"],
        SessionOperation.READ_CELLS,
        cells=[{"sheet": "Sheet1", "address": "C3"}],
    )

    first_descriptor = first["descriptor"]
    second_descriptor = second["descriptor"]
    assert first_descriptor["profile_identifier"] != second_descriptor["profile_identifier"]
    assert first_descriptor["uno_port"] != second_descriptor["uno_port"]
    assert first_descriptor["display"] != second_descriptor["display"]
    assert first_descriptor["working_copy"] is None
    assert (
        manager.descriptor(first["session_id"]).working_copy
        != manager.descriptor(second["session_id"]).working_copy
    )
    assert first_value["cells"][0]["value"] == 99
    assert second_value["cells"][0]["value"] is None


def test_failed_session_logs_survive_process_tree_cleanup(
    tmp_path: Path,
) -> None:
    backend = FakeSessionBackend(
        statuses={SessionOperation.RECALCULATE: GateExecutionStatus.FAILED}
    )
    manager = LibreOfficeSessionManager(
        backend=backend,
        workspace=WorkspacePathPolicy([tmp_path]),
        session_root=tmp_path / "sessions",
        id_factory=lambda: "failed-session",
    )
    source = tmp_path / "source.ods"
    source.write_bytes(b"original")
    session_id = manager.create_session()["session_id"]
    manager.open_document(session_id, source)

    failure = manager.perform(session_id, SessionOperation.RECALCULATE)
    destroyed = manager.destroy_session(session_id)
    logs = manager.collect_logs(session_id)

    assert failure["operation_status"] == "failed"
    assert destroyed["process_tree_cleaned"] is True
    assert destroyed["logs_preserved"] is True
    assert logs["state"] == "destroyed"
    assert logs["failed"] is True
    assert '"operation":"recalculate"' in logs["operation_log"]


def test_failed_creation_is_archived_with_diagnostics(tmp_path: Path) -> None:
    class UnavailableBackend(FakeSessionBackend):
        def create(
            self,
            *,
            session_id: str,
            profile_identifier: str,
            uno_port: int,
            display: str,
        ) -> SessionBackendResult:
            del session_id, profile_identifier, uno_port, display
            raise RuntimeError("runtime identity probe failed")

    backend = UnavailableBackend()
    manager = LibreOfficeSessionManager(
        backend=backend,
        workspace=WorkspacePathPolicy([tmp_path]),
        session_root=tmp_path / "sessions",
        id_factory=lambda: "creation-failure",
    )

    created = manager.create_session()
    logs = manager.collect_logs("creation-failure")

    assert created["operation_status"] == "unavailable"
    assert created["descriptor"]["state"] == "creation_failed"
    assert logs["failed"] is True
    assert '"operation":"create_session"' in logs["operation_log"]


def test_open_rejects_paths_outside_configured_workspace(
    manager: LibreOfficeSessionManager,
    tmp_path: Path,
) -> None:
    outside = tmp_path.parent / "outside-session.ods"
    outside.write_bytes(b"secret")
    session_id = manager.create_session()["session_id"]

    with pytest.raises(ValueError, match="outside configured workspace roots"):
        manager.open_document(session_id, outside)


def test_deterministic_in_process_mcp_client_requires_matching_session_ids(
    manager: LibreOfficeSessionManager,
    tmp_path: Path,
) -> None:
    configure_session_manager(manager)
    source = tmp_path / "source.ods"
    source.write_bytes(b"original")
    transport = InProcessSessionTransport(
        {
            "create_session": create_session,
            "open_document": open_document,
            "list_sheets": list_sheets,
            "close": close,
            "collect_logs": collect_logs,
            "destroy_session": destroy_session,
        }
    )
    client = LibreOfficeSessionClient(transport)

    async def exercise() -> tuple[dict[str, Any], dict[str, Any]]:
        created = await client.create_session()
        session_id = str(created["session_id"])
        await client.call(
            session_id,
            SessionOperation.OPEN_DOCUMENT,
            document_path=str(source),
        )
        sheets = await client.call(session_id, SessionOperation.LIST_SHEETS)
        await client.call(session_id, SessionOperation.CLOSE)
        await client.call(session_id, "destroy_session")
        logs = await client.call(session_id, "collect_logs")
        return sheets, logs

    try:
        sheets, logs = asyncio.run(exercise())
    finally:
        configure_session_manager(None)

    assert sheets["sheets"] == ["Sheet1"]
    assert logs["state"] == "destroyed"
    assert [tool for tool, _arguments in transport.calls] == [
        "create_session",
        "open_document",
        "list_sheets",
        "close",
        "destroy_session",
        "collect_logs",
    ]


def test_mcp_failure_response_preserves_requested_session_id(
    manager: LibreOfficeSessionManager,
) -> None:
    configure_session_manager(manager)
    try:
        response = asyncio.run(list_sheets("missing-session"))
    finally:
        configure_session_manager(None)

    assert response["session_id"] == "missing-session"
    assert response["transport_success"] is False
    assert response["operation_status"] == "failed"


def test_session_scenario_runner_uses_service_and_preserves_source(tmp_path: Path) -> None:
    source = tmp_path / "source.ods"
    source.write_bytes(b"source package")
    backend = FakeSessionBackend()
    manager = LibreOfficeSessionManager(
        backend=backend,
        workspace=WorkspacePathPolicy([tmp_path]),
        session_root=tmp_path / "sessions",
        id_factory=lambda: "acceptance-session",
        clock=lambda: datetime(2000, 1, 1, tzinfo=UTC),
    )
    scenario = Scenario(
        id="service-acceptance",
        steps=[
            ScenarioStep(
                id="open",
                action=Action(kind=ActionKind.OPEN),
            )
        ],
    )

    trace = LibreOfficeSessionScenarioRunner(manager=manager).run(
        source,
        EnvironmentManifest(),
        scenario,
    )
    logs = manager.collect_logs("acceptance-session")

    assert trace.status is GateExecutionStatus.PASSED
    assert trace.steps[0].step_id == "open"
    assert trace.runtime_identity.runtime_kind == "libreoffice_session_docker"
    assert source.read_bytes() == b"source package"
    assert logs["state"] == "destroyed"
    assert [operation for _session, operation, _arguments in backend.calls] == [
        "create_session",
        "open_document",
        "run_scenario",
        "destroy_session",
    ]


def test_session_scenario_runner_retains_failed_creation_evidence(tmp_path: Path) -> None:
    class UnavailableBackend(FakeSessionBackend):
        def create(
            self,
            *,
            session_id: str,
            profile_identifier: str,
            uno_port: int,
            display: str,
        ) -> SessionBackendResult:
            del session_id, profile_identifier, uno_port, display
            raise RuntimeError("pinned runtime unavailable")

    source = tmp_path / "source.ods"
    source.write_bytes(b"source package")
    manager = LibreOfficeSessionManager(
        backend=UnavailableBackend(),
        workspace=WorkspacePathPolicy([tmp_path]),
        session_root=tmp_path / "sessions",
        id_factory=lambda: "unavailable-session",
    )
    scenario = Scenario(
        id="creation-failure",
        steps=[ScenarioStep(id="open", action=Action(kind=ActionKind.OPEN))],
    )

    trace = LibreOfficeSessionScenarioRunner(manager=manager).run(
        source,
        EnvironmentManifest(),
        scenario,
    )

    assert trace.status is GateExecutionStatus.UNAVAILABLE
    assert trace.error is not None
    assert trace.error["type"] == "RuntimeError"
    assert any('"operation":"create_session"' in item for item in trace.logs)
    assert manager.collect_logs("unavailable-session")["state"] == "creation_failed"


def test_stable_session_service_cli_keeps_only_deprecated_legacy_alias() -> None:
    assert "libreoffice-mcp-serve" in cli.commands
    assert cli.commands["mcp-serve"].deprecated is True
