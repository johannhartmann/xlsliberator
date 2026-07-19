"""Office-free checks for the portable GUI replay artifact."""

import subprocess
from pathlib import Path
from typing import Any

import pytest

from xlsliberator.gui_worker import (
    _confined_path,
    _open_ready_document,
    _raise_with_office_diagnostics,
    _start_recording,
    _wait_for_startup_document,
    _write_replay_html,
)


def test_replay_html_embeds_timeline_without_interpreting_result_markup(tmp_path: Path) -> None:
    response = {
        "operations": [
            {
                "sequence": 1,
                "kind": "key",
                "status": "passed",
                "duration_ms": 125,
                "result": {"value": "</script><script>alert(1)</script>"},
            }
        ]
    }

    _write_replay_html(tmp_path, response)

    replay = (tmp_path / "replay.html").read_text(encoding="utf-8")
    assert 'src="recording.webm"' in replay
    assert "evidence.operations.forEach" in replay
    assert "</script><script>alert(1)</script>" not in replay
    assert "\\u003c/script>" in replay


def test_gui_worker_separates_read_only_inputs_from_job_outputs(tmp_path: Path) -> None:
    input_root = tmp_path / "input"
    job_root = tmp_path / "job"
    input_root.mkdir()
    job_root.mkdir()
    source = input_root / "target.ods"
    source.write_bytes(b"ods")

    assert (
        _confined_path(
            source,
            root=input_root,
            must_exist=True,
            label="input",
        )
        == source
    )
    assert (
        _confined_path(
            job_root / "evidence.zip",
            root=job_root,
            must_exist=False,
            label="output",
        )
        == job_root / "evidence.zip"
    )
    with pytest.raises(ValueError, match="must remain inside"):
        _confined_path(
            source,
            root=job_root,
            must_exist=True,
            label="output",
        )


def test_gui_document_waits_for_visible_calc_window_before_installing_controls(
    tmp_path: Path, monkeypatch: Any
) -> None:
    calls: list[str] = []
    document = object()

    def open_document(_session: dict[str, Any], _path: Path) -> object:
        calls.append("open")
        return document

    def wait_for_calc_window() -> str:
        calls.append("window")
        return "42"

    def install_controller(
        _session: dict[str, Any],
        received_document: object,
        _request: dict[str, Any],
    ) -> str:
        assert received_document is document
        calls.append("controls")
        return "controller"

    monkeypatch.setattr("xlsliberator.gui_worker._open_document", open_document)
    monkeypatch.setattr("xlsliberator.gui_worker._wait_for_calc_window", wait_for_calc_window)
    monkeypatch.setattr("xlsliberator.gui_worker._install_game_controller", install_controller)

    result = _open_ready_document({}, tmp_path / "target.ods", {})

    assert result == (document, "controller", "42")
    assert calls == ["open", "window", "controls"]


def test_gui_document_uses_native_startup_component_before_remote_reopen(
    tmp_path: Path, monkeypatch: Any
) -> None:
    calls: list[str] = []
    document = object()
    session: dict[str, Any] = {"startup_document_pending": True}

    monkeypatch.setattr(
        "xlsliberator.gui_worker._wait_for_startup_document",
        lambda _session: calls.append("startup") or document,
    )
    monkeypatch.setattr(
        "xlsliberator.gui_worker._open_document",
        lambda _session, _path: pytest.fail("startup document must not be opened twice"),
    )
    monkeypatch.setattr(
        "xlsliberator.gui_worker._wait_for_calc_window",
        lambda: calls.append("window") or "42",
    )
    monkeypatch.setattr(
        "xlsliberator.gui_worker._install_game_controller",
        lambda _session, _document, _request: calls.append("controls") or "controller",
    )

    result = _open_ready_document(session, tmp_path / "target.ods", {})

    assert result == (document, "controller", "42")
    assert session == {}
    assert calls == ["startup", "window", "controls"]


def test_startup_document_waits_for_calc_service(monkeypatch: Any) -> None:
    class Document:
        def supportsService(self, name: str) -> bool:  # noqa: N802
            return name == "com.sun.star.sheet.SpreadsheetDocument"

    class Desktop:
        def getCurrentComponent(self) -> Document:  # noqa: N802
            return Document()

    monkeypatch.setattr("xlsliberator.gui_worker._drain_ui", lambda _session: None)

    document = _wait_for_startup_document({"desktop": Desktop()})

    assert document.supportsService("com.sun.star.sheet.SpreadsheetDocument")


def test_gui_failure_preserves_bounded_desktop_diagnostics(
    tmp_path: Path, monkeypatch: Any
) -> None:
    xvfb_log = tmp_path / "xvfb.log"
    openbox_log = tmp_path / "openbox.log"
    xvfb_log.write_text("x" * 3_000, encoding="utf-8")
    openbox_log.write_text("openbox-ready", encoding="utf-8")
    monkeypatch.setenv("DISPLAY", ":99")
    monkeypatch.setenv("XLSLIBERATOR_XVFB_LOG", str(xvfb_log))
    monkeypatch.setenv("XLSLIBERATOR_OPENBOX_LOG", str(openbox_log))
    monkeypatch.setattr(
        "xlsliberator.gui_worker.subprocess.run",
        lambda *_args, **_kwargs: subprocess.CompletedProcess([], 0, "display-ready", ""),
    )
    monkeypatch.setattr(
        "xlsliberator.gui_worker._cgroup_memory_diagnostics",
        lambda: "cgroup_memory=memory.current=123, memory.max=456",
    )

    with pytest.raises(
        RuntimeError,
        match=r"(?s)bridge disposed.*cgroup_memory=memory.current=123, memory.max=456"
        r".*xvfb_log=x+.*openbox_log=openbox-ready"
        r".*x11_probe_exit_code=0.*display-ready",
    ):
        _raise_with_office_diagnostics(
            RuntimeError("bridge disposed"),
            {"office_exit_code": 255, "office_log": "office-abort"},
        )


def test_gui_recorder_uses_one_encoder_thread(tmp_path: Path, monkeypatch: Any) -> None:
    captured: list[list[str]] = []

    class Process:
        pass

    def popen(command: list[str], **_kwargs: Any) -> Process:
        captured.append(command)
        return Process()

    monkeypatch.setattr("xlsliberator.gui_worker.subprocess.Popen", popen)

    process = _start_recording(":99", tmp_path / "recording.webm")

    assert isinstance(process, Process)
    assert captured[0][captured[0].index("-threads") + 1] == "1"
