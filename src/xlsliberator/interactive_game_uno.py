"""Target-native LibreOffice adapter for the interactive-game showcase.

This module is imported by LibreOffice's bundled Python only.  It creates a
plain ODS target with native form controls and attaches bounded Python/UNO
listeners while the document is operated by the dedicated GUI runtime.  The
document contains no VBA project, LibreOffice Basic, or embedded script event
binding.
"""

from __future__ import annotations

import hashlib
import json
from contextlib import suppress
from pathlib import Path
from typing import Any, Final

from xlsliberator.interactive_game_engine import (
    GamePhase,
    GameState,
    move_left,
    move_right,
    new_game_state,
    pause_game,
    piece_cells,
    reset_game,
    resume_game,
    rotate_piece,
    soft_drop,
    start_game,
    state_from_json,
    state_to_json,
    tick,
    tick_interval_ms,
)

SOURCE_SHA256: Final = "da1bddc2c20ed8f5557b547e04a84cb1b476eca010e30a6be549be650894e4d1"
TARGET_BUILD: Final = "26.2.4.2"
GAME_SHEET: Final = "game"
SCORE_SHEET: Final = "Score"
STATE_SHEET: Final = "_XLSLIBERATOR_STATE"
STATE_CELL: Final = "A2"
BOARD_RANGE: Final = "M4:AA31"

CONTROL_NAMES: Final = (
    "GameStart",
    "GamePause",
    "GameReset",
    "GameHighScores",
    "ScoreReturn",
)

_BACKGROUND: Final = 0x101820
_GRID: Final = 0x1E293B
_ACTIVE: Final = 0x22C55E
_SETTLED_COLORS: Final[dict[int, int]] = {
    3: 0xF97316,
    4: 0xA855F7,
    6: 0x06B6D4,
    41: 0xEAB308,
    45: 0xEF4444,
    48: 0x3B82F6,
    54: 0xEC4899,
}


def build_interactive_game_target(request: dict[str, Any]) -> dict[str, Any]:
    """Create a script-free ODS application from the immutable real source."""
    from xlsliberator.lo_worker import (
        _close_document,
        _office_session,
        _property_value,
        _sha256_file,
    )

    source = Path(str(request["input_path"])).resolve()
    output = Path(str(request["output_path"])).resolve()
    if _sha256_file(source) != SOURCE_SHA256:
        raise ValueError("interactive-game source identity does not match the public corpus")
    if output.suffix.lower() != ".ods":
        raise ValueError("interactive-game target must use the ODS format")

    with _office_session(request, use_gui=False) as session:
        document = session["desktop"].loadComponentFromURL(
            "private:factory/scalc",
            "_blank",
            0,
            (_property_value("Hidden", True),),
        )
        if document is None:
            raise RuntimeError("LibreOffice did not create the interactive-game target")
        try:
            document.storeAsURL(
                session["uno"].systemPathToFileUrl(str(output)),
                (
                    _property_value("FilterName", "calc8"),
                    _property_value("Overwrite", True),
                ),
            )
            _initialize_document(document, session["uno"])
            _store_checkpoint(document, "final-render")
        except Exception:
            output.unlink(missing_ok=True)
            raise
        finally:
            _close_document(document, save=False)
    if not output.is_file():
        raise RuntimeError("LibreOffice did not write the interactive-game target")
    return {
        "status": "passed",
        "source_sha256": SOURCE_SHA256,
        "target_sha256": _sha256_file(output),
        "target_build": TARGET_BUILD,
        "target_format": "ods",
        "controls": list(CONTROL_NAMES),
        "state_sheet": STATE_SHEET,
        "board_range": BOARD_RANGE,
        "embedded_script_bindings": 0,
    }


def _initialize_document(document: Any, uno: Any) -> None:
    sheets = document.getSheets()
    first = sheets.getByIndex(0)
    first.setName(GAME_SHEET)
    sheets.insertNewByName(SCORE_SHEET, 1)
    game = sheets.getByName(GAME_SHEET)
    score = sheets.getByName(SCORE_SHEET)
    _add_form(document, game, score, uno)
    _store_checkpoint(document, "native-controls")

    sheets.insertNewByName(STATE_SHEET, 2)
    state_sheet = sheets.getByName(STATE_SHEET)

    _set_cell(game, "B2", "phase")
    _set_cell(game, "B3", "score")
    _set_cell(game, "B4", "completed lines")
    _set_cell(game, "B5", "high score")
    _set_cell(game, "B6", "active row")
    _set_cell(game, "B7", "event index")
    _set_cell(game, "B8", "timer ticks")
    _set_cell(game, "B9", "listener bindings")
    _set_cell(game, "B10", "active column")
    _set_cell(game, "B11", "rotation")
    _set_cell(game, "B13", "Keyboard: Left / Right / Down / Up or Ctrl")
    _set_cell(game, "B14", "Controls: Start / Pause / Reset / High Scores")
    _set_cell(game, "B15", "Sound: unavailable (no portable capability granted)")
    _set_cell(game, "M2", "XLSLiberator — LibreOffice Interactive Game")

    board = game.getCellRangeByName(BOARD_RANGE)
    board.CellBackColor = _BACKGROUND
    board.CharColor = 0xF8FAFC
    board.HoriJustify = 2
    for column in range(12, 27):
        game.getColumns().getByIndex(column).Width = 700
    for row in range(3, 31):
        game.getRows().getByIndex(row).Height = 500

    _set_cell(score, "A1", "High scores")
    _set_cell(score, "A3", "Player")
    _set_cell(score, "B3", "Score")
    _set_cell(score, "A4", "LibreOffice player")
    _set_cell(score, "B4", "0")

    _set_cell(state_sheet, "A1", "xlsliberator.interactive-game.state.v1")
    state_sheet.getCellRangeByName(STATE_CELL).setString(
        state_to_json(new_game_state(seed=20_260_204))
    )
    _set_cell(state_sheet, "A3", "0")
    _set_cell(state_sheet, "A4", SOURCE_SHA256)
    _set_cell(state_sheet, "A5", TARGET_BUILD)
    state_sheet.IsVisible = False

    _store_checkpoint(document, "base-document")
    controller = InteractiveGameController({}, document, enable_timer=False)
    try:
        controller.render()
    finally:
        controller.dispose()


def _set_cell(sheet: Any, address: str, value: str) -> None:
    sheet.getCellRangeByName(address).setString(value)


def _add_form(document: Any, game: Any, score: Any, uno: Any) -> None:
    game_draw_page = game.getDrawPage()
    game_forms = game_draw_page.getForms()
    game_form = document.createInstance("com.sun.star.form.component.DataForm")
    game_form.Name = "XLSLiberatorGameControls"
    game_forms.insertByIndex(game_forms.getCount(), game_form)
    for index, (name, label) in enumerate(
        (
            ("GameStart", "Start"),
            ("GamePause", "Pause / Resume"),
            ("GameReset", "Reset"),
            ("GameHighScores", "High Scores"),
        )
    ):
        _add_button(
            document,
            game,
            game_form,
            uno,
            name=name,
            label=label,
            x=1_000,
            y=4_500 + index * 1_250,
            width=4_000,
        )

    score_draw_page = score.getDrawPage()
    score_forms = score_draw_page.getForms()
    score_form = document.createInstance("com.sun.star.form.component.DataForm")
    score_form.Name = "XLSLiberatorScoreControls"
    score_forms.insertByIndex(score_forms.getCount(), score_form)
    _add_button(
        document,
        score,
        score_form,
        uno,
        name="ScoreReturn",
        label="Return to Game",
        x=1_000,
        y=5_000,
        width=4_000,
    )


def _add_button(
    document: Any,
    sheet: Any,
    form: Any,
    uno: Any,
    *,
    name: str,
    label: str,
    x: int,
    y: int,
    width: int,
) -> None:
    model = document.createInstance("com.sun.star.form.component.CommandButton")
    model.Name = name
    model.Label = label
    shape = document.createInstance("com.sun.star.drawing.ControlShape")
    position = uno.createUnoStruct("com.sun.star.awt.Point")
    position.X = x
    position.Y = y
    size = uno.createUnoStruct("com.sun.star.awt.Size")
    size.Width = width
    size.Height = 900
    shape.setPosition(position)
    shape.setSize(size)
    shape.setControl(model)
    form.insertByIndex(form.getCount(), model)
    sheet.getDrawPage().add(shape)


def _store_checkpoint(document: Any, stage: str) -> None:
    try:
        document.store()
    except Exception as exc:
        raise RuntimeError(f"interactive-game store checkpoint failed: {stage}") from exc


class InteractiveGameController:
    """Bridge Calc controller events to immutable game-state transitions."""

    def __init__(
        self,
        session: dict[str, Any],
        document: Any,
        *,
        enable_timer: bool = True,
    ) -> None:
        self.session = session
        self.document = document
        self.enable_timer = enable_timer
        self.state = self._load_state()
        self.listeners: list[tuple[Any, Any]] = []
        self.key_listener: Any = None
        self.timer_listener: Any = None
        self.timer: Any = None
        self.disposed = False
        self.event_log: list[dict[str, Any]] = []

    def install(self) -> None:
        if self.disposed:
            raise RuntimeError("cannot install a disposed game controller")
        if self.listeners or self.key_listener is not None:
            raise RuntimeError("game controller listeners are already installed")
        controller = self.document.getCurrentController()
        if hasattr(controller, "setFormDesignMode"):
            controller.setFormDesignMode(False)
        action_type, key_type, timer_type = _listener_types()
        for name in CONTROL_NAMES:
            model = _find_control_model(self.document, name)
            listener = action_type(self)
            model.addActionListener(listener)
            self.listeners.append((model, listener))
        self.key_listener = key_type(self)
        controller.addKeyHandler(self.key_listener)

        if self.enable_timer:
            context = self.session["component_context"]
            manager = context.getServiceManager()
            self.timer = manager.createInstanceWithContext("com.sun.star.awt.Timer", context)
            self.timer_listener = timer_type(self)
            self.timer.addTimerListener(self.timer_listener)
            self._sync_timer()
        self._record("document-open")
        self._increment_open_count()
        self.persist_and_render()

    def dispose(self) -> None:
        if self.disposed:
            return
        self.disposed = True
        if self.timer is not None:
            with suppress(Exception):
                self.timer.stop()
            if self.timer_listener is not None:
                with suppress(Exception):
                    self.timer.removeTimerListener(self.timer_listener)
        controller = None
        with suppress(Exception):
            controller = self.document.getCurrentController()
        if controller is not None and self.key_listener is not None:
            with suppress(Exception):
                controller.removeKeyHandler(self.key_listener)
        for model, listener in self.listeners:
            with suppress(Exception):
                model.removeActionListener(listener)
        self.listeners.clear()
        self._record("document-close")

    def action(self, control_name: str) -> None:
        before = self.state
        if control_name == "GameStart":
            self.state = start_game(self.state)
        elif control_name == "GamePause":
            if self.state.phase is GamePhase.RUNNING:
                self.state = pause_game(self.state)
            elif self.state.phase is GamePhase.PAUSED:
                self.state = resume_game(self.state)
        elif control_name == "GameReset":
            self.state = reset_game(self.state)
        elif control_name == "GameHighScores":
            self.document.getCurrentController().setActiveSheet(
                self.document.getSheets().getByName(SCORE_SHEET)
            )
        elif control_name == "ScoreReturn":
            self.document.getCurrentController().setActiveSheet(
                self.document.getSheets().getByName(GAME_SHEET)
            )
        else:
            raise ValueError(f"unknown game control: {control_name}")
        self._record("control", control_name=control_name, changed=self.state != before)
        self.persist_and_render()

    def key(self, key_code: int, modifiers: int) -> bool:
        keys = _key_constants()
        before = self.state
        if key_code == keys["LEFT"]:
            self.state = move_left(self.state)
        elif key_code == keys["RIGHT"]:
            self.state = move_right(self.state)
        elif key_code == keys["DOWN"]:
            self.state = soft_drop(self.state)
        elif key_code == keys["UP"] or (
            modifiers & keys["MOD1"] and key_code in {keys["CONTROL"], 0}
        ):
            self.state = rotate_piece(self.state)
        elif key_code == keys["ESCAPE"]:
            self.state = (
                pause_game(self.state)
                if self.state.phase is GamePhase.RUNNING
                else resume_game(self.state)
                if self.state.phase is GamePhase.PAUSED
                else self.state
            )
        else:
            return False
        self._record(
            "keyboard",
            key_code=key_code,
            modifiers=modifiers,
            changed=self.state != before,
        )
        self.persist_and_render()
        return True

    def timer_tick(self) -> None:
        before = self.state
        self.state = tick(self.state)
        self._record("timer", changed=self.state != before)
        self.persist_and_render()

    def load_fixture(self, payload: str) -> None:
        self.state = state_from_json(payload)
        self._record("public-fixture")
        self.persist_and_render()

    def persist_and_render(self) -> None:
        state_sheet = self.document.getSheets().getByName(STATE_SHEET)
        state_sheet.getCellRangeByName(STATE_CELL).setString(state_to_json(self.state))
        self.render()
        self._sync_timer()

    def render(self) -> None:
        sheets = self.document.getSheets()
        game = sheets.getByName(GAME_SHEET)
        board = game.getCellRangeByName(BOARD_RANGE)
        board.clearContents(1023)
        board.CellBackColor = _GRID
        for cell in self.state.settled:
            game.getCellByPosition(
                12 + cell.column, 3 + cell.row
            ).CellBackColor = _SETTLED_COLORS.get(cell.color, 0x94A3B8)
        if self.state.active is not None:
            for row, column in piece_cells(self.state.active):
                game.getCellByPosition(12 + column, 3 + row).CellBackColor = _ACTIVE

        _set_cell(game, "C2", self.state.phase.value)
        _set_cell(game, "C3", str(self.state.score))
        _set_cell(game, "C4", str(self.state.completed_lines))
        _set_cell(game, "C5", str(self.state.high_score))
        _set_cell(
            game,
            "C6",
            str(self.state.active.row) if self.state.active is not None else "",
        )
        _set_cell(game, "C7", str(self.state.event_index))
        _set_cell(game, "C8", str(self.state.tick_index))
        _set_cell(game, "C9", str(len(self.listeners) + (self.key_listener is not None)))
        _set_cell(
            game,
            "C10",
            str(self.state.active.column) if self.state.active is not None else "",
        )
        _set_cell(
            game,
            "C11",
            str(self.state.active.rotation) if self.state.active is not None else "",
        )
        _set_cell(sheets.getByName(SCORE_SHEET), "B4", str(self.state.high_score))

    def evidence(self) -> dict[str, Any]:
        encoded = state_to_json(self.state)
        return {
            "state": json.loads(encoded),
            "state_sha256": hashlib.sha256(encoded.encode()).hexdigest(),
            "events": list(self.event_log),
            "control_bindings": len(self.listeners),
            "key_handler_installed": self.key_listener is not None,
            "timer_installed": self.timer_listener is not None,
        }

    def _load_state(self) -> GameState:
        sheets = self.document.getSheets()
        if not sheets.hasByName(STATE_SHEET):
            raise ValueError("interactive-game target is missing its state sheet")
        encoded = sheets.getByName(STATE_SHEET).getCellRangeByName(STATE_CELL).getString()
        return state_from_json(str(encoded))

    def _increment_open_count(self) -> None:
        cell = self.document.getSheets().getByName(STATE_SHEET).getCellRangeByName("A3")
        value = int(str(cell.getString()) or "0") + 1
        cell.setString(str(value))

    def _sync_timer(self) -> None:
        if self.timer is None:
            return
        with suppress(Exception):
            self.timer.stop()
        if self.state.phase is GamePhase.RUNNING:
            self.timer.setTimeout(tick_interval_ms(self.state))
            self.timer.start()

    def _record(self, kind: str, **details: Any) -> None:
        self.event_log.append(
            {
                "sequence": len(self.event_log) + 1,
                "kind": kind,
                **details,
            }
        )


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
    raise ValueError(f"interactive-game native control is missing: {name}")


def _listener_types() -> tuple[type[Any], type[Any], type[Any]]:
    import unohelper
    from com.sun.star.awt import XActionListener, XKeyHandler, XTimerListener

    class ActionListener(unohelper.Base, XActionListener):
        def __init__(self, adapter: InteractiveGameController) -> None:
            self.adapter = adapter

        def actionPerformed(self, event: Any) -> None:  # noqa: N802
            model = getattr(event.Source, "Model", None)
            name = str(getattr(model, "Name", ""))
            self.adapter.action(name)

        def disposing(self, _event: Any) -> None:
            return None

    class KeyListener(unohelper.Base, XKeyHandler):
        def __init__(self, adapter: InteractiveGameController) -> None:
            self.adapter = adapter

        def keyPressed(self, event: Any) -> bool:  # noqa: N802
            return self.adapter.key(int(event.KeyCode), int(event.Modifiers))

        def keyReleased(self, _event: Any) -> bool:  # noqa: N802
            return False

        def disposing(self, _event: Any) -> None:
            return None

    class TimerListener(unohelper.Base, XTimerListener):
        def __init__(self, adapter: InteractiveGameController) -> None:
            self.adapter = adapter

        def timeout(self, _event: Any) -> None:
            self.adapter.timer_tick()

        def disposing(self, _event: Any) -> None:
            return None

    return ActionListener, KeyListener, TimerListener


def _key_constants() -> dict[str, int]:
    from com.sun.star.awt import Key, KeyModifier

    return {
        "LEFT": int(Key.LEFT),
        "RIGHT": int(Key.RIGHT),
        "DOWN": int(Key.DOWN),
        "UP": int(Key.UP),
        "ESCAPE": int(Key.ESCAPE),
        "CONTROL": int(getattr(Key, "CONTROL", 0)),
        "MOD1": int(KeyModifier.MOD1),
    }


__all__ = [
    "BOARD_RANGE",
    "CONTROL_NAMES",
    "GAME_SHEET",
    "InteractiveGameController",
    "SOURCE_SHA256",
    "STATE_CELL",
    "STATE_SHEET",
    "TARGET_BUILD",
    "build_interactive_game_target",
]
