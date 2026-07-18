"""Real Docker-only integration tests for the stateful LibreOffice service."""

from __future__ import annotations

import subprocess
from pathlib import Path

import pytest

from xlsliberator.embed_macros import embed_python_macros
from xlsliberator.execution_sandbox import WorkspacePathPolicy
from xlsliberator.libreoffice_session import (
    DockerSessionBackend,
    LibreOfficeSessionManager,
    SessionOperation,
)
from xlsliberator.scenarios.models import EnvironmentManifest


@pytest.mark.integration
@pytest.mark.docker
def test_stateful_service_open_read_write_macro_save_close_reopen_and_export(
    tmp_path: Path,
    skip_if_no_lo: None,
) -> None:
    source = tmp_path / "source.ods"
    source.write_bytes(
        (Path(__file__).parents[1] / "fixtures" / "scenarios" / "basic.ods").read_bytes()
    )
    embed_python_macros(
        source,
        {
            "session_macro.py": (
                "def set_value():\n"
                "    document = XSCRIPTCONTEXT.getDocument()\n"
                "    document.getSheets().getByName('Sheet1').getCellRangeByName('A1').setValue(11)\n"
                "    return 11\n"
            )
        },
    )
    before = source.read_bytes()
    environment = EnvironmentManifest(
        declared_capabilities=["macro_execution"],
        granted_capabilities=["macro_execution"],
    )
    manager = LibreOfficeSessionManager(
        backend=DockerSessionBackend(),
        workspace=WorkspacePathPolicy([tmp_path]),
        session_root=tmp_path / "sessions",
        id_factory=lambda: "real-session",
    )

    created = manager.create_session(environment=environment)
    session_id = str(created["session_id"])
    opened = manager.open_document(session_id, source)
    sheets = manager.perform(session_id, SessionOperation.LIST_SHEETS)
    written = manager.perform(
        session_id,
        SessionOperation.WRITE_CELLS,
        cells=[{"sheet": "Sheet1", "address": "A1", "value": 7}],
    )
    recalculated = manager.perform(session_id, SessionOperation.RECALCULATE)
    formula = manager.perform(
        session_id,
        SessionOperation.READ_CELLS,
        cells=[{"sheet": "Sheet1", "address": "A3"}],
    )
    macro = manager.perform(
        session_id,
        SessionOperation.EXECUTE_PYTHON_MACRO,
        script_uri=(
            "vnd.sun.star.script:session_macro.py$set_value?language=Python&location=document"
        ),
    )
    macro_value = manager.perform(
        session_id,
        SessionOperation.READ_CELLS,
        cells=[{"sheet": "Sheet1", "address": "A1"}],
    )
    controls = manager.perform(
        session_id,
        SessionOperation.DISPATCH_CONTROL_EVENT,
        control_name="Missing",
    )
    saved_path = tmp_path / "saved.ods"
    saved = manager.perform(
        session_id,
        SessionOperation.SAVE,
        output_path=str(saved_path),
    )
    closed = manager.perform(session_id, SessionOperation.CLOSE)
    reopened = manager.perform(session_id, SessionOperation.REOPEN)
    reopened_value = manager.perform(
        session_id,
        SessionOperation.READ_CELLS,
        cells=[{"sheet": "Sheet1", "address": "A1"}],
    )
    pdf_path = tmp_path / "export.pdf"
    exported = manager.perform(
        session_id,
        SessionOperation.EXPORT_PDF,
        output_path=str(pdf_path),
    )
    destroyed = manager.destroy_session(session_id)
    logs = manager.collect_logs(session_id)

    assert created["success"] is True, created
    assert created["descriptor"]["libreoffice_version"] == "26.2.4.2"
    assert opened["success"] is True, opened
    assert sheets["sheets"] == ["Sheet1"]
    assert written["success"] is True, written
    assert recalculated["success"] is True, recalculated
    assert formula["cells"][0]["value"] == 10
    assert macro["success"] is True, macro
    assert macro_value["cells"][0]["value"] == 11
    assert controls["operation_status"] == "unavailable"
    assert controls["implemented"] is False
    assert saved["success"] is True
    assert saved_path.is_file()
    assert closed["session_state"] == "closed"
    assert reopened["session_state"] == "open"
    assert reopened_value["cells"][0]["value"] == 11
    assert exported["success"] is True, exported
    assert pdf_path.read_bytes().startswith(b"%PDF")
    assert destroyed["process_tree_cleaned"] is True
    assert logs["state"] == "destroyed"
    assert logs["operation_log"]
    assert source.read_bytes() == before


@pytest.mark.integration
@pytest.mark.docker
def test_stateful_service_timeout_cleans_the_named_container_process_tree(
    tmp_path: Path,
    skip_if_no_lo: None,
) -> None:
    source = tmp_path / "source.ods"
    source.write_bytes(
        (Path(__file__).parents[1] / "fixtures" / "scenarios" / "basic.ods").read_bytes()
    )
    backend = DockerSessionBackend()
    manager = LibreOfficeSessionManager(
        backend=backend,
        workspace=WorkspacePathPolicy([tmp_path]),
        session_root=tmp_path / "sessions",
        id_factory=lambda: "timeout-session",
    )
    session_id = str(manager.create_session()["session_id"])
    opened = manager.open_document(session_id, source)
    assert opened["success"] is True, opened
    before = _office_containers()
    backend.timeout_seconds = 0.001

    timed_out = manager.perform(session_id, SessionOperation.LIST_SHEETS)
    after = _office_containers()
    manager.destroy_session(session_id)

    assert timed_out["operation_status"] == "unavailable"
    assert timed_out["transport_success"] is False
    assert after == before


def _office_containers() -> set[str]:
    result = subprocess.run(  # nosec B603 - fixed Docker-only diagnostic command
        ["docker", "ps", "--filter", "name=xlsliberator-lo-", "--format", "{{.Names}}"],
        check=True,
        capture_output=True,
        text=True,
        timeout=15,
    )
    return {line for line in result.stdout.splitlines() if line}
