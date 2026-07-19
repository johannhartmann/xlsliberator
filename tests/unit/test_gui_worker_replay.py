"""Office-free checks for the portable GUI replay artifact."""

import subprocess
from pathlib import Path
from typing import Any

import pytest

from xlsliberator.gui_worker import (
    _confined_path,
    _open_ready_document,
    _raise_with_office_diagnostics,
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

    with pytest.raises(
        RuntimeError,
        match=r"(?s)bridge disposed.*xvfb_log=x+.*openbox_log=openbox-ready"
        r".*x11_probe_exit_code=0.*display-ready",
    ):
        _raise_with_office_diagnostics(
            RuntimeError("bridge disposed"),
            {"office_exit_code": 255, "office_log": "office-abort"},
        )
