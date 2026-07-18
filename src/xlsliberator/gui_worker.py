"""Bounded real-X11 interaction runner for the dedicated office GUI container."""

from __future__ import annotations

import hashlib
import json
import os
import re
import subprocess
import time
from pathlib import Path
from typing import Any
from zipfile import ZIP_DEFLATED, ZipFile

_ALLOWED_KEYS = frozenset({"Left", "Right", "Down", "Up", "ctrl", "Escape", "space"})
_ALLOWED_ACTIONS = frozenset(
    {
        "click_control",
        "key",
        "wait",
        "observe",
        "recalculate",
        "save",
        "close",
        "reopen",
        "screenshot",
        "load_game_state",
    }
)
_SAFE_NAME = re.compile(r"^[A-Za-z0-9_.-]{1,100}$")


def run_gui_scenario(request: dict[str, Any]) -> dict[str, Any]:
    """Execute declared X11 events and lifecycle actions in one office process."""
    _require_gui_container()
    from xlsliberator.lo_worker import (
        _close_document,
        _office_session,
        _sha256_file,
    )

    source = _job_path(request["ods_path"], must_exist=True)
    archive_path = _job_path(request["output_path"], must_exist=False)
    if archive_path.suffix.lower() != ".zip":
        raise ValueError("GUI scenario output must be a ZIP evidence bundle")
    output_dir = archive_path.parent / ".gui-evidence"
    output_dir.mkdir(parents=True, exist_ok=True)
    actions = request.get("actions")
    if not isinstance(actions, list) or not 1 <= len(actions) <= 100:
        raise ValueError("run_gui_scenario requires between 1 and 100 actions")

    display = str(os.environ.get("DISPLAY") or "")
    if not re.fullmatch(r":\d{1,5}", display):
        raise RuntimeError("dedicated GUI runtime requires a private X11 display")

    working_copy = output_dir / "working-copy.ods"
    working_copy.write_bytes(source.read_bytes())
    recording = output_dir / "recording.webm"
    video = _start_recording(display, recording)
    records: list[dict[str, Any]] = []
    screenshots: list[str] = []
    controller_evidence: list[dict[str, Any]] = []

    try:
        with _office_session(
            {
                **request,
                "session_display": display,
            },
            use_gui=True,
        ) as session:
            document = _open_document(session, working_copy)
            game_controller = _install_game_controller(session, document, request)
            window_id = _wait_for_calc_window()
            try:
                for sequence, raw_action in enumerate(actions, start=1):
                    if not isinstance(raw_action, dict):
                        raise ValueError("GUI actions must be objects")
                    kind = str(raw_action.get("kind") or "")
                    if kind not in _ALLOWED_ACTIONS:
                        raise ValueError(f"unsupported GUI action: {kind}")
                    started = time.monotonic()
                    before = _document_state_hash(document, raw_action)
                    result: dict[str, Any] = {}

                    if kind == "click_control":
                        name = _safe_name(raw_action.get("control_name"), "control_name")
                        _click_control(document, name, window_id)
                        result = {"control_name": name, "event_surface": "x11-pointer"}
                    elif kind == "key":
                        key = str(raw_action.get("key") or "")
                        if key not in _ALLOWED_KEYS:
                            raise ValueError(f"unsupported GUI key: {key}")
                        _xdotool("key", "--clearmodifiers", "--window", window_id, key)
                        result = {"key": key, "event_surface": "x11-keyboard"}
                    elif kind == "wait":
                        seconds = float(raw_action.get("seconds", 0))
                        if not 0 <= seconds <= 5:
                            raise ValueError("GUI wait must be between zero and five seconds")
                        time.sleep(seconds)
                        result = {"seconds": seconds}
                    elif kind == "observe":
                        result = _observe(document, raw_action)
                    elif kind == "recalculate":
                        document.calculateAll()
                        result = {"recalculated": True}
                    elif kind == "save":
                        document.store()
                        result = {"saved_sha256": _sha256_file(working_copy)}
                    elif kind == "close":
                        if game_controller is not None:
                            controller_evidence.append(game_controller.evidence())
                            game_controller.dispose()
                            game_controller = None
                        _close_document(document, save=False)
                        document = None
                        result = {"closed": True}
                    elif kind == "reopen":
                        if game_controller is not None:
                            controller_evidence.append(game_controller.evidence())
                            game_controller.dispose()
                        if document is not None:
                            _close_document(document, save=False)
                        document = _open_document(session, working_copy)
                        game_controller = _install_game_controller(session, document, request)
                        window_id = _wait_for_calc_window()
                        result = {"reopened_sha256": _sha256_file(working_copy)}
                    elif kind == "screenshot":
                        name = _safe_name(raw_action.get("name"), "screenshot name")
                        screenshot = output_dir / f"{name}.png"
                        _run_checked(["scrot", "--focused", str(screenshot)])
                        screenshots.append(screenshot.name)
                        result = {"path": screenshot.name, "sha256": _sha256_file(screenshot)}
                    elif kind == "load_game_state":
                        if game_controller is None:
                            raise RuntimeError(
                                "load_game_state requires the interactive-game adapter"
                            )
                        state_json = str(raw_action.get("state_json") or "")
                        if len(state_json.encode()) > 100_000:
                            raise ValueError("game fixture exceeds the bounded state size")
                        game_controller.load_fixture(state_json)
                        result = {
                            "loaded_state_sha256": hashlib.sha256(state_json.encode()).hexdigest()
                        }

                    _drain_ui(session)
                    after = _document_state_hash(document, raw_action)
                    records.append(
                        {
                            "sequence": sequence,
                            "kind": kind,
                            "status": "passed",
                            "duration_ms": round((time.monotonic() - started) * 1000, 3),
                            "state_sha256_before": before,
                            "state_sha256_after": after,
                            "result": result,
                        }
                    )
            finally:
                if game_controller is not None:
                    controller_evidence.append(game_controller.evidence())
                    game_controller.dispose()
                _close_document(document, save=False)
    finally:
        _stop_recording(video)

    if not recording.is_file() or recording.stat().st_size == 0:
        raise RuntimeError("GUI recording was not produced")
    response = {
        "status": "passed",
        "event_layer": "xvfb-openbox-xdotool",
        "display": display,
        "source_sha256": _sha256_file(source),
        "working_copy_sha256": _sha256_file(working_copy),
        "recording": {
            "path": recording.name,
            "sha256": _sha256_file(recording),
            "bytes": recording.stat().st_size,
        },
        "screenshots": screenshots,
        "operations": records,
        "controller_sessions": controller_evidence,
    }
    (output_dir / "result.json").write_text(
        json.dumps(response, indent=2, sort_keys=True) + "\n",
        encoding="utf-8",
    )
    with ZipFile(archive_path, "w", compression=ZIP_DEFLATED) as archive:
        for path in sorted(output_dir.rglob("*")):
            if path.is_file():
                archive.write(path, path.relative_to(output_dir).as_posix())
    if not archive_path.is_file() or archive_path.stat().st_size == 0:
        raise RuntimeError("GUI evidence archive was not produced")
    response["evidence_archive_sha256"] = _sha256_file(archive_path)
    response["evidence_archive_bytes"] = archive_path.stat().st_size
    return response


def _require_gui_container() -> None:
    if (
        os.environ.get("XLSLIBERATOR_OFFICE_CONTAINER") != "1"
        or os.environ.get("XLSLIBERATOR_UI_EVENT_CONTAINER") != "1"
        or not Path("/.dockerenv").is_file()
    ):
        raise RuntimeError("real UI events require the dedicated pinned Docker GUI runtime")


def _job_path(raw: object, *, must_exist: bool) -> Path:
    root = Path(os.environ.get("XLSLIBERATOR_JOB_DIR", "/job")).resolve()
    path = Path(str(raw)).resolve()
    if path != root and root not in path.parents:
        raise ValueError("GUI scenario paths must remain inside the container job directory")
    if must_exist and not path.is_file():
        raise FileNotFoundError(path)
    return path


def _safe_name(raw: object, label: str) -> str:
    value = str(raw or "")
    if not _SAFE_NAME.fullmatch(value):
        raise ValueError(f"{label} is malformed")
    return value


def _open_document(session: dict[str, Any], path: Path) -> Any:
    from xlsliberator.lo_worker import _property_value

    document = session["desktop"].loadComponentFromURL(
        session["uno"].systemPathToFileUrl(str(path)),
        "_blank",
        0,
        (
            _property_value("Hidden", False),
            _property_value("MacroExecutionMode", 4),
        ),
    )
    if document is None:
        raise RuntimeError("LibreOffice did not open the GUI scenario document")
    controller = document.getCurrentController()
    if hasattr(controller, "setFormDesignMode"):
        controller.setFormDesignMode(False)
    _drain_ui(session)
    return document


def _install_game_controller(
    session: dict[str, Any],
    document: Any,
    request: dict[str, Any],
) -> Any:
    if request.get("adapter") != "interactive-game":
        return None
    from xlsliberator.interactive_game_uno import InteractiveGameController

    controller = InteractiveGameController(
        session,
        document,
        enable_timer=bool(request.get("timer_enabled", True)),
    )
    controller.install()
    _drain_ui(session)
    return controller


def _wait_for_calc_window() -> str:
    deadline = time.monotonic() + 15
    while time.monotonic() < deadline:
        result = subprocess.run(
            ["xdotool", "search", "--onlyvisible", "--class", "libreoffice-calc"],
            capture_output=True,
            text=True,
            check=False,
        )
        windows = [line.strip() for line in result.stdout.splitlines() if line.strip().isdigit()]
        if windows:
            window_id = windows[-1]
            _xdotool("windowactivate", "--sync", window_id)
            return window_id
        time.sleep(0.1)
    raise RuntimeError("LibreOffice Calc X11 window did not become visible")


def _click_control(document: Any, name: str, window_id: str) -> None:
    model = _find_control_model(document, name)
    controller = document.getCurrentController()
    view = controller.getControl(model)
    if view is None:
        raise RuntimeError(f"control view is unavailable: {name}")
    position = view.getPosSize()
    geometry = _window_geometry(window_id)
    x = geometry["X"] + int(position.X) + max(1, int(position.Width) // 2)
    y = geometry["Y"] + int(position.Y) + max(1, int(position.Height) // 2)
    _xdotool("mousemove", "--sync", str(x), str(y))
    _xdotool("click", "1")


def _find_control_model(document: Any, name: str) -> Any:
    sheets = document.getSheets()
    for sheet_index in range(sheets.getCount()):
        forms = sheets.getByIndex(sheet_index).getDrawPage().getForms()
        for form_index in range(forms.getCount()):
            form = forms.getByIndex(form_index)
            for control_index in range(form.getCount()):
                control = form.getByIndex(control_index)
                if getattr(control, "Name", None) == name:
                    return control
    raise ValueError(f"control was not found: {name}")


def _window_geometry(window_id: str) -> dict[str, int]:
    result = _run_checked(["xdotool", "getwindowgeometry", "--shell", window_id])
    values: dict[str, int] = {}
    for line in result.stdout.splitlines():
        key, separator, value = line.partition("=")
        if separator and key in {"X", "Y", "WIDTH", "HEIGHT"}:
            values[key] = int(value)
    if set(values) != {"X", "Y", "WIDTH", "HEIGHT"}:
        raise RuntimeError("xdotool returned incomplete window geometry")
    return values


def _observe(document: Any, action: dict[str, Any]) -> dict[str, Any]:
    sheet_name = _safe_name(action.get("sheet"), "sheet")
    address = str(action.get("address") or "")
    if not re.fullmatch(r"[A-Za-z]{1,3}[1-9][0-9]{0,5}", address):
        raise ValueError("observation address is malformed")
    sheets = document.getSheets()
    if not sheets.hasByName(sheet_name):
        raise ValueError(f"observation sheet is missing: {sheet_name}")
    cell = sheets.getByName(sheet_name).getCellRangeByName(address)
    result = {
        "sheet": sheet_name,
        "address": address,
        "string": str(cell.getString()),
        "value": float(cell.getValue()),
        "formula": str(cell.getFormula()),
        "background": int(cell.CellBackColor),
    }
    expected_string = action.get("expect_string")
    if expected_string is not None and result["string"] != str(expected_string):
        raise AssertionError(
            f"{sheet_name}!{address} string is {result['string']!r}, "
            f"expected {expected_string!r}"
        )
    expected_value = action.get("expect_value")
    if expected_value is not None and result["value"] != float(expected_value):
        raise AssertionError(
            f"{sheet_name}!{address} value is {result['value']!r}, "
            f"expected {expected_value!r}"
        )
    minimum = action.get("min_value")
    if minimum is not None and result["value"] < float(minimum):
        raise AssertionError(
            f"{sheet_name}!{address} value is below {minimum!r}: {result['value']!r}"
        )
    maximum = action.get("max_value")
    if maximum is not None and result["value"] > float(maximum):
        raise AssertionError(
            f"{sheet_name}!{address} value is above {maximum!r}: {result['value']!r}"
        )
    expected_background = action.get("expect_background")
    if (
        expected_background is not None
        and result["background"] != int(expected_background)
    ):
        raise AssertionError(
            f"{sheet_name}!{address} background is {result['background']!r}, "
            f"expected {expected_background!r}"
        )
    return result


def _document_state_hash(document: Any, action: dict[str, Any]) -> str:
    if document is None:
        return hashlib.sha256(b"closed").hexdigest()
    observations = action.get("state_cells")
    if not isinstance(observations, list):
        observations = [
            {"sheet": "_XLSLIBERATOR_STATE", "address": "A2"},
            {"sheet": "game", "address": "B2"},
            {"sheet": "game", "address": "B3"},
        ]
    values: list[dict[str, Any]] = []
    for observation in observations[:20]:
        if not isinstance(observation, dict):
            raise ValueError("state_cells entries must be objects")
        values.append(_observe(document, observation))
    encoded = json.dumps(values, sort_keys=True, separators=(",", ":")).encode()
    return hashlib.sha256(encoded).hexdigest()


def _drain_ui(session: dict[str, Any]) -> None:
    del session
    time.sleep(0.15)


def _xdotool(*arguments: str) -> None:
    _run_checked(["xdotool", *arguments])


def _run_checked(command: list[str]) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        command,
        capture_output=True,
        text=True,
        timeout=20,
        check=True,
    )


def _start_recording(display: str, path: Path) -> subprocess.Popen[bytes]:
    return subprocess.Popen(
        [
            "ffmpeg",
            "-nostdin",
            "-loglevel",
            "error",
            "-f",
            "x11grab",
            "-video_size",
            "1280x1024",
            "-framerate",
            "15",
            "-i",
            f"{display}.0",
            "-c:v",
            "libvpx-vp9",
            "-deadline",
            "realtime",
            "-y",
            str(path),
        ],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.PIPE,
    )


def _stop_recording(process: subprocess.Popen[bytes]) -> None:
    if process.poll() is None:
        process.terminate()
    try:
        _, stderr = process.communicate(timeout=10)
    except subprocess.TimeoutExpired:
        process.kill()
        _, stderr = process.communicate(timeout=5)
    if process.returncode not in {0, 255}:
        raise RuntimeError(f"GUI recording failed: {stderr.decode(errors='replace')[-1000:]}")
