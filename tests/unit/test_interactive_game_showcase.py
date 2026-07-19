"""Office-free contract tests for the Docker-only interactive-game boundary."""

from __future__ import annotations

import hashlib
import json
from pathlib import Path
from typing import Any

import pytest

from xlsliberator.interactive_game_engine import new_game_state, start_game, state_from_json
from xlsliberator.interactive_game_showcase import (
    GUI_IMAGE,
    PUBLIC_SCENARIOS,
    _require_success,
    build_target,
    bundle_gui_replays,
    run_gui_scenario,
)
from xlsliberator.interactive_game_uno import (
    CONTROL_NAMES,
    SOURCE_SHA256,
    InteractiveGameController,
    _control_logical_name,
    _set_numeric_cell,
    _set_optional_numeric_cell,
)


class _FakeRuntime:
    instances: list[_FakeRuntime] = []

    def __init__(self, image: str | None = None, **_kwargs: Any) -> None:
        self.image = image
        self.payload: dict[str, Any] | None = None
        self.instances.append(self)

    def request(self, payload: dict[str, Any]) -> dict[str, Any]:
        self.payload = payload
        return {
            "success": True,
            "data": {
                "status": "passed",
                "target_build": "26.2.4.2",
            },
        }


def test_build_target_uses_only_the_pinned_worker_boundary(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    source = tmp_path / "source.xlsb"
    source.write_bytes(b"source")
    output = tmp_path / "target.ods"
    monkeypatch.setattr(
        "xlsliberator.interactive_game_showcase.LibreOfficeDockerRuntime",
        _FakeRuntime,
    )
    _FakeRuntime.instances.clear()

    result = build_target(source, output)

    assert result["status"] == "passed"
    assert len(_FakeRuntime.instances) == 1
    runtime = _FakeRuntime.instances[0]
    assert runtime.image is None
    assert runtime.payload == {
        "op": "build_interactive_game_target",
        "input_path": str(source),
        "output_path": str(output),
        "timeout_seconds": 120,
    }


def test_gui_scenario_selects_dedicated_image_and_forwards_timer_policy(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    target = tmp_path / "target.ods"
    target.write_bytes(b"target")
    output = tmp_path / "evidence.zip"
    actions = [{"kind": "observe", "sheet": "game", "address": "C2"}]
    monkeypatch.setattr(
        "xlsliberator.interactive_game_showcase.LibreOfficeDockerRuntime",
        _FakeRuntime,
    )
    _FakeRuntime.instances.clear()

    run_gui_scenario(target, output, actions, timer_enabled=False)

    runtime = _FakeRuntime.instances[0]
    assert runtime.image == GUI_IMAGE
    assert runtime.payload is not None
    assert runtime.payload["op"] == "run_gui_scenario"
    assert runtime.payload["adapter"] == "interactive-game"
    assert runtime.payload["timer_enabled"] is False
    assert runtime.payload["actions"] == actions


def test_failed_gui_scenario_preserves_worker_and_container_diagnostics() -> None:
    with pytest.raises(
        RuntimeError,
        match=r"(?s)bridge disposed.*office_exit_code=134.*container stderr:.*outer stderr",
    ):
        _require_success(
            {
                "success": False,
                "data": {"container_stderr": "outer stderr"},
                "error": {
                    "message": "bridge disposed",
                    "traceback": "office_exit_code=134",
                },
            },
            "interactive-game GUI scenario",
        )


def test_replay_bundle_selects_gui_image_and_requires_all_public_scenarios(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    evidence: dict[str, Path] = {}
    for scenario_id in PUBLIC_SCENARIOS:
        path = tmp_path / f"{scenario_id}.zip"
        path.write_bytes(scenario_id.encode())
        evidence[scenario_id] = path
    output = tmp_path / "replay.zip"
    monkeypatch.setattr(
        "xlsliberator.interactive_game_showcase.LibreOfficeDockerRuntime",
        _FakeRuntime,
    )
    _FakeRuntime.instances.clear()

    bundle_gui_replays(evidence, output)

    runtime = _FakeRuntime.instances[0]
    assert runtime.image == GUI_IMAGE
    assert runtime.payload is not None
    assert runtime.payload["op"] == "bundle_gui_replays"
    assert runtime.payload["output_path"] == str(output)
    assert not Path(runtime.payload["input_path"]).exists()

    with pytest.raises(ValueError, match="every canonical public scenario"):
        bundle_gui_replays(
            {scenario: path for scenario, path in evidence.items() if scenario != "timer-tick"},
            output,
        )


def test_public_scenarios_have_exact_ids_and_valid_game_fixtures() -> None:
    root = Path("demos/interactive-game/showcase/scenarios")
    scenarios = [
        json.loads(path.read_text(encoding="utf-8")) for path in sorted(root.glob("*.json"))
    ]

    assert {scenario["scenario_id"] for scenario in scenarios} == {
        "keyboard-control",
        "timer-tick",
        "native-controls",
        "document-events",
        "line-collapse",
    }
    for scenario in scenarios:
        assert scenario["schema_version"] == "1.0.0"
        assert scenario["actions"]
        for action in scenario["actions"]:
            if action["kind"] == "load_game_state":
                state_from_json(action["state_json"])


def test_source_identity_is_bound_to_the_real_public_workbook() -> None:
    source = Path("demos/interactive-game/source/TetrisGameDemo.xlsb")

    assert source.is_file()
    assert hashlib.sha256(source.read_bytes()).hexdigest() == SOURCE_SHA256


def test_native_control_logical_name_prefers_tag_and_supports_legacy_name() -> None:
    tagged = type("Control", (), {"Name": "Control1", "Tag": "GameStart"})()
    legacy = type("Control", (), {"Name": "GameStart", "Tag": ""})()
    native = type("Control", (), {"Name": "CertificationButton", "Tag": ""})()

    assert _control_logical_name(tagged) == "GameStart"
    assert _control_logical_name(legacy) == "GameStart"
    assert _control_logical_name(native) == "GameStart"


def test_external_timer_pump_emits_due_source_cadence_without_uno_timer(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr("xlsliberator.interactive_game_uno.time.monotonic", lambda: 10.7)
    controller = object.__new__(InteractiveGameController)
    controller.enable_timer = True
    controller.disposed = False
    controller.state = start_game(new_game_state(seed=17))
    controller._timer_last_poll = 10.0
    controller._timer_budget_ms = 0.0
    ticks: list[int] = []

    def tick_once() -> None:
        ticks.append(len(ticks) + 1)

    monkeypatch.setattr(controller, "timer_tick", tick_once)

    emitted = controller.pump_timer()

    assert emitted == 4
    assert ticks == [1, 2, 3, 4]
    assert controller._timer_budget_ms == pytest.approx(60.0)


def test_action_listeners_bind_to_native_control_views(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    class Listener:
        def __init__(self, adapter: InteractiveGameController) -> None:
            self.adapter = adapter

    class ControlView:
        def __init__(self) -> None:
            self.listeners: list[Listener] = []
            self.design_mode = True

        def addActionListener(self, listener: Listener) -> None:  # noqa: N802
            self.listeners.append(listener)

        def removeActionListener(self, listener: Listener) -> None:  # noqa: N802
            self.listeners.remove(listener)

        def setDesignMode(self, enabled: bool) -> None:  # noqa: N802
            self.design_mode = enabled

        def isDesignMode(self) -> bool:  # noqa: N802
            return self.design_mode

        def setEnable(self, _enabled: bool) -> None:  # noqa: N802
            return None

        def setVisible(self, _visible: bool) -> None:  # noqa: N802
            return None

    models = {name: object() for name in CONTROL_NAMES}
    views = {model: ControlView() for model in models.values()}
    game_sheet = object()
    score_sheet = object()
    active_sheets: list[object] = [game_sheet]

    def set_active_sheet(_self: object, sheet: object) -> None:
        active_sheets.append(sheet)

    native_controller = type(
        "NativeController",
        (),
        {
            "getActiveSheet": lambda _self: active_sheets[-1],
            "getControl": lambda _self, model: views[model],
            "setActiveSheet": set_active_sheet,
        },
    )()
    sheets = type(
        "Sheets",
        (),
        {"getByName": lambda _self, name: score_sheet if name == "Score" else game_sheet},
    )()
    document = type(
        "Document",
        (),
        {
            "getCurrentController": lambda _self: native_controller,
            "getSheets": lambda _self: sheets,
        },
    )()
    controller = object.__new__(InteractiveGameController)
    controller.document = document
    controller.listeners = []
    monkeypatch.setattr(
        "xlsliberator.interactive_game_uno._listener_types",
        lambda: (Listener, object),
    )
    monkeypatch.setattr(
        "xlsliberator.interactive_game_uno._find_control_model",
        lambda _document, name: models[name],
    )

    controller._attach_action_listeners()

    assert [control for _name, control, _listener in controller.listeners] == list(views.values())
    assert all(len(view.listeners) == 1 for view in views.values())
    assert all(not view.design_mode for view in views.values())
    assert active_sheets[-2:] == [score_sheet, game_sheet]

    replacement = ControlView()
    controller.ensure_action_listener("GameStart", replacement)

    assert len(controller.listeners) == len(CONTROL_NAMES)
    assert views[models["GameStart"]].listeners == []
    assert len(replacement.listeners) == 1


def test_game_metrics_use_calc_numeric_cells() -> None:
    class Cell:
        def __init__(self) -> None:
            self.numeric: float | None = None
            self.text: str | None = None

        def setValue(self, value: float) -> None:  # noqa: N802
            self.numeric = value

        def setString(self, value: str) -> None:  # noqa: N802
            self.text = value

    cell = Cell()
    sheet = type("Sheet", (), {"getCellRangeByName": lambda _self, _address: cell})()

    _set_numeric_cell(sheet, "C10", 5)

    assert cell.numeric == 5.0
    assert cell.text is None

    _set_optional_numeric_cell(sheet, "C10", None)

    assert cell.text == ""
