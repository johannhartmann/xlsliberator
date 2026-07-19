"""Office-free tests for the source-derived interactive-game candidate."""

from __future__ import annotations

import hashlib
import json
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path("demos/interactive-game/candidate").resolve()))

from candidate_interactive_game.adapter import (  # noqa: E402
    CONTROL_NAMES,
    SOURCE_SHA256,
    InteractiveGameController,
    _control_logical_name,
    _set_numeric_cell,
    _set_optional_numeric_cell,
)
from candidate_interactive_game.engine import (  # noqa: E402
    new_game_state,
    start_game,
    state_from_json,
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
            if action["kind"] == "load_fixture":
                state_from_json(action["payload"])


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
    monkeypatch.setattr("candidate_interactive_game.adapter.time.monotonic", lambda: 10.7)
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
        "candidate_interactive_game.adapter._listener_types",
        lambda: (Listener, object),
    )
    monkeypatch.setattr(
        "candidate_interactive_game.adapter._find_control_model",
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
