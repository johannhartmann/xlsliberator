"""Office-free checks for the portable GUI replay artifact."""

import subprocess
from pathlib import Path
from typing import Any

import pytest

from xlsliberator.gui_worker import (
    _cleanup_gui_session,
    _click_control,
    _concatenate_recordings,
    _confined_path,
    _control_screen_rectangle,
    _open_ready_document,
    _raise_with_office_diagnostics,
    _safe_key,
    _start_recording,
    _wait_for_startup_document,
    _write_replay_html,
)


def test_gui_key_contract_is_generic_but_rejects_xdotool_injection() -> None:
    assert _safe_key("Left") == "Left"
    assert _safe_key("ctrl+s") == "ctrl+s"
    assert _safe_key("F12") == "F12"
    assert _safe_key("A") == "A"

    for unsafe in ("", "ctrl+shift+alt+s+x", "Left --window 1", "XF86Launch1", "F13"):
        with pytest.raises(ValueError, match="unsupported GUI key"):
            _safe_key(unsafe)


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
        received_factory: Any,
        _request: dict[str, Any],
    ) -> str:
        assert received_document is document
        assert received_factory is factory
        calls.append("controls")
        return "controller"

    factory = object()
    monkeypatch.setattr("xlsliberator.gui_worker._open_document", open_document)
    monkeypatch.setattr("xlsliberator.gui_worker._wait_for_calc_window", wait_for_calc_window)
    monkeypatch.setattr(
        "xlsliberator.gui_worker._install_application_controller",
        install_controller,
    )

    result = _open_ready_document({}, tmp_path / "target.ods", factory, {})

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
        "xlsliberator.gui_worker._install_application_controller",
        lambda _session, _document, _factory, _request: calls.append("controls") or "controller",
    )

    result = _open_ready_document(session, tmp_path / "target.ods", object(), {})

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


def test_gui_failure_preserves_x11_error_before_abort_marker(
    tmp_path: Path, monkeypatch: Any
) -> None:
    monkeypatch.setenv("DISPLAY", ":99")
    monkeypatch.setattr(
        "xlsliberator.gui_worker.subprocess.run",
        lambda *_args, **_kwargs: subprocess.CompletedProcess([], 0, "display-ready", ""),
    )
    monkeypatch.setattr(
        "xlsliberator.gui_worker._cgroup_memory_diagnostics",
        lambda: "cgroup_memory=memory.current=123",
    )
    preamble = "X-Error: BadAlloc\n\tMajor opcode: 53 (X_CreatePixmap)\n"
    marker = "XLSLIBERATOR_ABORT_BACKTRACE_BEGIN\nframe\n"

    with pytest.raises(
        RuntimeError,
        match=r"(?s)X-Error: BadAlloc.*X_CreatePixmap"
        r".*XLSLIBERATOR_ABORT_BACKTRACE_BEGIN.*frame",
    ):
        _raise_with_office_diagnostics(
            RuntimeError("bridge disposed"),
            {"office_exit_code": 134, "office_log": preamble + marker},
        )


def test_gui_cleanup_preserves_primary_failure_but_not_standalone_cleanup_error() -> None:
    evidence: list[dict[str, Any]] = []

    class Controller:
        def evidence(self) -> dict[str, Any]:
            return {"status": "partial"}

        def dispose(self) -> None:
            raise RuntimeError("disposed bridge")

    def close_document(_document: Any, *, save: bool) -> None:
        del save
        raise RuntimeError("closed bridge")

    _cleanup_gui_session(
        Controller(),
        object(),
        evidence,
        close_document,
        preserve_primary_error=True,
    )

    assert evidence == [{"status": "partial"}]
    with pytest.raises(RuntimeError, match="disposed bridge"):
        _cleanup_gui_session(
            Controller(),
            object(),
            [],
            close_document,
            preserve_primary_error=False,
        )


def test_native_control_geometry_uses_accessible_screen_coordinates() -> None:
    class Geometry:
        X = 70
        Y = 171
        Width = 198
        Height = 45

    class Context:
        def getLocationOnScreen(self) -> Geometry:  # noqa: N802
            return Geometry()

        def getSize(self) -> Geometry:  # noqa: N802
            return Geometry()

    class View:
        def getAccessibleContext(self) -> Context:  # noqa: N802
            return Context()

        def getPosSize(self) -> None:  # noqa: N802
            pytest.fail("parent-relative XWindow geometry must not drive X11 clicks")

    assert _control_screen_rectangle(View(), "GameStart") == {
        "X": 70,
        "Y": 171,
        "WIDTH": 198,
        "HEIGHT": 45,
    }


def test_native_pointer_click_requires_matching_uno_action(monkeypatch: Any) -> None:
    events: list[dict[str, Any]] = []

    class Sheet:
        Name = "game"

    class Controller:
        def getActiveSheet(self) -> Sheet:  # noqa: N802
            return Sheet()

    class Document:
        def getCurrentController(self) -> Controller:  # noqa: N802
            return Controller()

    class GameController:
        def evidence(self) -> dict[str, Any]:
            return {"events": list(events)}

        def ensure_action_listener(self, name: str, view: Any) -> None:
            assert name == "GameStart"
            assert view is control_view

    control_view = object()
    pointer_events: list[tuple[str, ...]] = []

    def xdotool(*arguments: str) -> None:
        pointer_events.append(arguments)
        if arguments == ("mouseup", "1"):
            events.append(
                {
                    "kind": "control",
                    "control_name": "GameStart",
                    "sequence": 2,
                }
            )

    monkeypatch.setattr("xlsliberator.gui_worker._find_control_model", lambda *_args: object())
    monkeypatch.setattr(
        "xlsliberator.gui_worker._wait_for_control_view",
        lambda *_args: control_view,
    )
    monkeypatch.setattr(
        "xlsliberator.gui_worker._control_screen_rectangle",
        lambda *_args: {"X": 70, "Y": 171, "WIDTH": 198, "HEIGHT": 45},
    )
    monkeypatch.setattr(
        "xlsliberator.gui_worker._window_geometry",
        lambda *_args: {"X": 0, "Y": 0, "WIDTH": 1280, "HEIGHT": 1024},
    )
    monkeypatch.setattr("xlsliberator.gui_worker._xdotool", xdotool)
    monkeypatch.setattr("xlsliberator.gui_worker._drain_ui", lambda _session: None)

    assert _click_control({}, Document(), GameController(), "GameStart", "42") == {
        "control_name": "GameStart",
        "event_surface": "x11-pointer",
        "event_sequence": 2,
        "screen_rectangle": {"X": 70, "Y": 171, "WIDTH": 198, "HEIGHT": 45},
    }
    assert pointer_events[-3:] == [
        ("mousemove", "--sync", "169", "193"),
        ("mousedown", "1"),
        ("mouseup", "1"),
    ]


def test_recording_concat_manifest_never_enters_public_replay(
    tmp_path: Path, monkeypatch: Any
) -> None:
    recording = tmp_path / "scenario.webm"
    recording.write_bytes(b"\x1aE\xdf\xa3scenario")
    replay_dir = tmp_path / "public" / "replay"
    replay_dir.mkdir(parents=True)
    output = replay_dir / "showcase.webm"

    def run(command: list[str], **_kwargs: Any) -> subprocess.CompletedProcess[str]:
        Path(command[-1]).write_bytes(b"\x1aE\xdf\xa3combined")
        return subprocess.CompletedProcess(command, 0, "", "")

    monkeypatch.setattr("xlsliberator.gui_worker.subprocess.run", run)

    _concatenate_recordings([recording], output)

    assert output.read_bytes() == b"\x1aE\xdf\xa3combined"
    assert not (replay_dir / "recordings.txt").exists()


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
