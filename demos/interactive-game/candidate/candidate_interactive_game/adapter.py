"""Target-native LibreOffice adapter generated for the interactive-game episode.

This module is imported by LibreOffice's bundled Python only.  It creates a
plain ODS target and installs native form controls only inside the dedicated
Docker GUI runtime.  Controls are detached before Calc saves because pinned
LibreOffice 26.2.4.2 crashes while exporting CommandButton models.  The saved
document contains no VBA project, LibreOffice Basic, embedded script event
binding, or persisted form model.
"""

from __future__ import annotations

import hashlib
import json
import time
from contextlib import suppress
from pathlib import Path
from typing import Any, Final

from candidate_interactive_game.engine import (
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
_LOGICAL_CONTROL_NAMES: Final = {
    "CertificationButton": "GameStart",
    "Control2": "GamePause",
    "Control3": "GameReset",
    "Control4": "GameHighScores",
    "ScoreReturnButton": "ScoreReturn",
}
_RUNTIME_FORM_NAMES: Final = (
    "XLSLiberatorGameControls",
    "XLSLiberatorScoreControls",
)
_GAME_CONTROLS: Final = (
    ("GameStart", "Start"),
    ("GamePause", "Pause / Resume"),
    ("GameReset", "Reset"),
    ("GameHighScores", "High Scores"),
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


def build_target(request: dict[str, Any]) -> dict[str, Any]:
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

    output.unlink(missing_ok=True)
    try:
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
                _create_game_sheets(document)
                game = _prepare_game_sheet(document)
                _initialize_document(document, game)
                document.storeAsURL(
                    session["uno"].systemPathToFileUrl(str(output)),
                    (
                        _property_value("FilterName", "calc8"),
                        _property_value("Overwrite", True),
                    ),
                )
            except Exception:
                output.unlink(missing_ok=True)
                raise
            finally:
                _close_document(document, save=False)
            _verify_clean_target(session, output, _property_value, _close_document)
    except Exception:
        output.unlink(missing_ok=True)
        raise
    if not output.is_file():
        raise RuntimeError("LibreOffice did not write the interactive-game target")
    return {
        "status": "passed",
        "source_sha256": SOURCE_SHA256,
        "target_sha256": _sha256_file(output),
        "target_build": TARGET_BUILD,
        "target_format": "ods",
        "controls": list(CONTROL_NAMES),
        "control_lifecycle": "docker-runtime-native",
        "state_sheet": STATE_SHEET,
        "board_range": BOARD_RANGE,
        "embedded_script_bindings": 0,
    }


def create_controller(
    session: dict[str, Any],
    document: Any,
    config: dict[str, Any],
) -> InteractiveGameController:
    """Create the source-derived live controller through the generic runtime contract."""
    return InteractiveGameController(
        session,
        document,
        enable_timer=bool(config.get("timer_enabled", True)),
    )


def _prepare_game_sheet(document: Any) -> Any:
    sheets = document.getSheets()
    game = sheets.getByName(GAME_SHEET)
    _set_cell(game, "A1", "XLSLiberator interactive game target")
    return game


def _create_game_sheets(document: Any) -> None:
    sheets = document.getSheets()
    first = sheets.getByIndex(0)
    first.setName(GAME_SHEET)
    sheets.insertNewByName(SCORE_SHEET, 1)
    sheets.insertNewByName(STATE_SHEET, 2)


def _verify_clean_target(
    session: dict[str, Any],
    output: Path,
    property_value: Any,
    close_document: Any,
) -> None:
    output_url = session["uno"].systemPathToFileUrl(str(output))
    document = session["desktop"].loadComponentFromURL(
        output_url,
        "_blank",
        0,
        (property_value("Hidden", True),),
    )
    if document is None:
        raise RuntimeError("LibreOffice did not reopen the interactive-game target")
    try:
        state_sheet = document.getSheets().getByName(STATE_SHEET)
        persisted_names = {
            str(state_sheet.getCellRangeByName(f"A{row}").getString())
            for row in range(7, 7 + len(CONTROL_NAMES))
        }
        if persisted_names != set(CONTROL_NAMES):
            raise RuntimeError("interactive-game target lost its runtime control manifest")
        _assert_no_persisted_form_controls(document)
    finally:
        close_document(document, save=False)


def _initialize_document(document: Any, game: Any) -> None:
    sheets = document.getSheets()
    score = sheets.getByName(SCORE_SHEET)
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
    _set_numeric_cell(score, "B4", 0)

    _set_cell(state_sheet, "A1", "xlsliberator.interactive-game.state.v1")
    state_sheet.getCellRangeByName(STATE_CELL).setString(
        state_to_json(new_game_state(seed=20_260_204))
    )
    _set_numeric_cell(state_sheet, "A3", 0)
    _set_cell(state_sheet, "A4", SOURCE_SHA256)
    _set_cell(state_sheet, "A5", TARGET_BUILD)
    for row, control_name in enumerate(CONTROL_NAMES, start=7):
        _set_cell(state_sheet, f"A{row}", control_name)
    state_sheet.IsVisible = False

    controller = InteractiveGameController({}, document, enable_timer=False)
    try:
        controller.render()
    finally:
        controller.dispose()


def _set_cell(sheet: Any, address: str, value: str) -> None:
    sheet.getCellRangeByName(address).setString(value)


def _set_numeric_cell(sheet: Any, address: str, value: int) -> None:
    """Write a real Calc number so public assertions can use XCell.getValue."""
    sheet.getCellRangeByName(address).setValue(float(value))


def _set_optional_numeric_cell(
    sheet: Any,
    address: str,
    value: int | None,
) -> None:
    if value is None:
        _set_cell(sheet, address, "")
    else:
        _set_numeric_cell(sheet, address, value)


def _install_runtime_controls(document: Any, uno: Any) -> None:
    """Create native UNO controls only for the live Docker-contained session."""
    _assert_no_persisted_form_controls(document)
    sheets = document.getSheets()
    game = sheets.getByName(GAME_SHEET)
    score = sheets.getByName(SCORE_SHEET)
    game_form = document.createInstance("com.sun.star.form.component.Form")
    game_form.Name = _RUNTIME_FORM_NAMES[0]
    game.getDrawPage().getForms().insertByName(game_form.Name, game_form)
    for index, (name, label) in enumerate(_GAME_CONTROLS):
        _add_runtime_button(
            document,
            game,
            game_form,
            uno,
            name=name,
            label=label,
            x=1_000,
            y=1_000 + index * 1_500,
            width=5_000,
        )

    score_form = document.createInstance("com.sun.star.form.component.Form")
    score_form.Name = _RUNTIME_FORM_NAMES[1]
    score.getDrawPage().getForms().insertByName(score_form.Name, score_form)
    _add_runtime_button(
        document,
        score,
        score_form,
        uno,
        name="ScoreReturn",
        label="Return to Game",
        x=1_000,
        y=1_000,
        width=5_000,
    )


def _add_runtime_button(
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
    model.Tag = name
    model.Label = label
    model.Tabstop = True
    form.insertByName(name, model)
    shape = document.createInstance("com.sun.star.drawing.ControlShape")
    position = uno.createUnoStruct("com.sun.star.awt.Point")
    position.X = x
    position.Y = y
    size = uno.createUnoStruct("com.sun.star.awt.Size")
    size.Width = width
    size.Height = 1_200
    shape.setPosition(position)
    shape.setSize(size)
    shape.setControl(model)
    sheet.getDrawPage().add(shape)


def _remove_runtime_controls(document: Any) -> None:
    sheets = document.getSheets()
    for sheet_index in range(sheets.getCount()):
        draw_page = sheets.getByIndex(sheet_index).getDrawPage()
        for shape_index in reversed(range(draw_page.getCount())):
            shape = draw_page.getByIndex(shape_index)
            model = None
            with suppress(Exception):
                model = shape.getControl()
            if model is not None and _control_logical_name(model) in CONTROL_NAMES:
                draw_page.remove(shape)
        forms = draw_page.getForms()
        for form_name in _RUNTIME_FORM_NAMES:
            if forms.hasByName(form_name):
                forms.removeByName(form_name)
        for shape_index in range(draw_page.getCount()):
            shape = draw_page.getByIndex(shape_index)
            model = None
            with suppress(Exception):
                model = shape.getControl()
            if model is not None and _control_logical_name(model) in CONTROL_NAMES:
                raise RuntimeError("LibreOffice retained a transient native control shape")


def _assert_no_persisted_form_controls(document: Any) -> None:
    sheets = document.getSheets()
    for sheet_index in range(sheets.getCount()):
        if sheets.getByIndex(sheet_index).getDrawPage().getForms().getCount():
            raise RuntimeError("interactive-game controls must not enter the ODS form exporter")


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
        self.listeners: list[tuple[str, Any, Any]] = []
        self.key_listener: Any = None
        self._timer_last_poll = time.monotonic()
        self._timer_budget_ms = 0.0
        self.disposed = False
        self.runtime_controls_installed = False
        self.event_log: list[dict[str, Any]] = []

    def install(self) -> None:
        if self.disposed:
            raise RuntimeError("cannot install a disposed game controller")
        if self.listeners or self.key_listener is not None:
            raise RuntimeError("game controller listeners are already installed")
        controller = self.document.getCurrentController()
        _install_live_runtime_controls(self.document, self.session["uno"], controller)
        self.runtime_controls_installed = True
        _action_type, key_type = _listener_types()
        self._attach_action_listeners()
        self.key_listener = key_type(self)
        controller.addKeyHandler(self.key_listener)

        self._record("document-open")
        self._increment_open_count()
        self.persist_and_render()
        self._reset_timer_clock()

    def dispose(self) -> None:
        if self.disposed:
            return
        self.disposed = True
        self._timer_budget_ms = 0.0
        controller = None
        with suppress(Exception):
            controller = self.document.getCurrentController()
        if controller is not None and self.key_listener is not None:
            with suppress(Exception):
                controller.removeKeyHandler(self.key_listener)
        self._detach_action_listeners()
        if self.runtime_controls_installed:
            _remove_runtime_controls(self.document)
            self.runtime_controls_installed = False
        self._record("document-close")

    def prepare_for_save(self) -> None:
        """Detach transient controls before LibreOffice enters its form exporter."""
        if self.disposed or not self.runtime_controls_installed:
            raise RuntimeError("game controller has no live runtime controls")
        self._detach_action_listeners()
        _remove_runtime_controls(self.document)
        self.runtime_controls_installed = False

    def restore_after_save(self) -> None:
        """Restore target-native controls after a clean ODS save."""
        if self.disposed or self.runtime_controls_installed:
            raise RuntimeError("game controller cannot restore runtime controls")
        controller = self.document.getCurrentController()
        _install_live_runtime_controls(self.document, self.session["uno"], controller)
        self.runtime_controls_installed = True
        self._attach_action_listeners()

    def _attach_action_listeners(self) -> None:
        if self.listeners:
            raise RuntimeError("runtime control listeners are already attached")
        controller = self.document.getCurrentController()
        sheets = self.document.getSheets()
        original_sheet = controller.getActiveSheet()
        try:
            for name in CONTROL_NAMES:
                sheet_name = SCORE_SHEET if name == "ScoreReturn" else GAME_SHEET
                controller.setActiveSheet(sheets.getByName(sheet_name))
                model = _find_control_model(self.document, name)
                control = _wait_for_control_view(controller, model, name)
                self.ensure_action_listener(name, control)
        finally:
            controller.setActiveSheet(original_sheet)

    def ensure_action_listener(self, name: str, control: Any) -> None:
        """Bind the adapter to the currently materialized native control view."""
        retained: list[tuple[str, Any, Any]] = []
        for bound_name, bound_control, listener in self.listeners:
            if bound_name == name:
                with suppress(Exception):
                    bound_control.removeActionListener(listener)
            else:
                retained.append((bound_name, bound_control, listener))
        action_type, _key_type = _listener_types()
        listener = action_type(self)
        control.addActionListener(listener)
        retained.append((name, control, listener))
        self.listeners = retained

    def _detach_action_listeners(self) -> None:
        for _name, control, listener in self.listeners:
            with suppress(Exception):
                control.removeActionListener(listener)
        self.listeners.clear()

    def action(self, control_name: str) -> None:
        before = self.state
        previous_phase = before.phase
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
        if self.state.phase is not previous_phase:
            self._reset_timer_clock()
        self._record("control", control_name=control_name, changed=self.state != before)
        self.persist_and_render()

    def key(self, key_code: int, modifiers: int) -> bool:
        keys = _key_constants()
        before = self.state
        previous_phase = before.phase
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
        if self.state.phase is not previous_phase:
            self._reset_timer_clock()
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

    def pump_timer(self) -> int:
        """Advance due ticks from the bounded Docker worker's monotonic clock.

        LibreOffice's published AWT UNO API has no timer service or timer
        listener interface.  The external PyUNO controller therefore owns the
        source-derived cadence and invokes state transitions on its single
        action thread after real X11/UNO events have drained.
        """
        now = time.monotonic()
        elapsed_ms = max(0.0, (now - self._timer_last_poll) * 1_000)
        self._timer_last_poll = now
        if not self.enable_timer or self.disposed or self.state.phase is not GamePhase.RUNNING:
            self._timer_budget_ms = 0.0
            return 0

        # GUI actions are limited to five seconds.  The cap prevents a paused
        # or externally stalled worker from replaying an unbounded tick burst.
        self._timer_budget_ms += min(elapsed_ms, 10_000.0)
        emitted = 0
        while emitted < 256 and self.state.phase is GamePhase.RUNNING:
            interval_ms = tick_interval_ms(self.state)
            if self._timer_budget_ms < interval_ms:
                break
            self._timer_budget_ms -= interval_ms
            self.timer_tick()
            emitted += 1
        return emitted

    def load_fixture(self, payload: str) -> None:
        self.state = state_from_json(payload)
        self._reset_timer_clock()
        self._record("public-fixture")
        self.persist_and_render()

    def persist_and_render(self) -> None:
        state_sheet = self.document.getSheets().getByName(STATE_SHEET)
        state_sheet.getCellRangeByName(STATE_CELL).setString(state_to_json(self.state))
        self.render()

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
        _set_numeric_cell(game, "C3", self.state.score)
        _set_numeric_cell(game, "C4", self.state.completed_lines)
        _set_numeric_cell(game, "C5", self.state.high_score)
        _set_optional_numeric_cell(
            game,
            "C6",
            self.state.active.row if self.state.active is not None else None,
        )
        _set_numeric_cell(game, "C7", self.state.event_index)
        _set_numeric_cell(game, "C8", self.state.tick_index)
        _set_numeric_cell(game, "C9", len(self.listeners) + (self.key_listener is not None))
        _set_optional_numeric_cell(
            game,
            "C10",
            self.state.active.column if self.state.active is not None else None,
        )
        _set_optional_numeric_cell(
            game,
            "C11",
            self.state.active.rotation if self.state.active is not None else None,
        )
        _set_numeric_cell(sheets.getByName(SCORE_SHEET), "B4", self.state.high_score)

    def evidence(self) -> dict[str, Any]:
        encoded = state_to_json(self.state)
        return {
            "state": json.loads(encoded),
            "state_sha256": hashlib.sha256(encoded.encode()).hexdigest(),
            "events": list(self.event_log),
            "control_bindings": len(self.listeners),
            "key_handler_installed": self.key_listener is not None,
            "timer_installed": self.enable_timer,
            "timer_surface": "docker-worker-monotonic-pump",
        }

    def _load_state(self) -> GameState:
        sheets = self.document.getSheets()
        if not sheets.hasByName(STATE_SHEET):
            raise ValueError("interactive-game target is missing its state sheet")
        encoded = sheets.getByName(STATE_SHEET).getCellRangeByName(STATE_CELL).getString()
        return state_from_json(str(encoded))

    def _increment_open_count(self) -> None:
        cell = self.document.getSheets().getByName(STATE_SHEET).getCellRangeByName("A3")
        cell.setValue(float(int(cell.getValue()) + 1))

    def _reset_timer_clock(self) -> None:
        self._timer_last_poll = time.monotonic()
        self._timer_budget_ms = 0.0

    def _record(self, kind: str, **details: Any) -> None:
        self.event_log.append(
            {
                "sequence": len(self.event_log) + 1,
                "kind": kind,
                **details,
            }
        )


def _install_live_runtime_controls(document: Any, uno: Any, controller: Any) -> None:
    """Materialize transient form controls in the current Calc view."""
    if not hasattr(controller, "setFormDesignMode") or not hasattr(controller, "isFormDesignMode"):
        raise RuntimeError("Calc controller does not expose the native form-layer lifecycle")
    controller.setFormDesignMode(True)
    try:
        _install_runtime_controls(document, uno)
    finally:
        controller.setFormDesignMode(False)
    if bool(controller.isFormDesignMode()):
        raise RuntimeError("Calc form layer remained in design mode")


def _wait_for_control_view(controller: Any, model: Any, name: str) -> Any:
    deadline = time.monotonic() + 2.0
    last_error: Exception | None = None
    while time.monotonic() < deadline:
        try:
            control = controller.getControl(model)
            if control is not None:
                control.setDesignMode(False)
                control.setEnable(True)
                control.setVisible(True)
                if not bool(control.isDesignMode()):
                    return control
                last_error = RuntimeError(f"native control remained in design mode: {name}")
        except Exception as exc:
            last_error = exc
        time.sleep(0.05)
    raise RuntimeError(f"native control view is unavailable: {name}") from last_error


def _find_control_model(document: Any, name: str) -> Any:
    sheets = document.getSheets()
    for sheet_index in range(sheets.getCount()):
        forms = sheets.getByIndex(sheet_index).getDrawPage().getForms()
        for form_index in range(forms.getCount()):
            form = forms.getByIndex(form_index)
            for control_index in range(form.getCount()):
                control = form.getByIndex(control_index)
                if _control_logical_name(control) == name:
                    return control
    raise ValueError(f"interactive-game native control is missing: {name}")


def _control_logical_name(model: Any) -> str:
    """Return the public control ID without relying on fragile UNO model names."""
    tag = str(getattr(model, "Tag", "") or "")
    if tag:
        return tag
    native_name = str(getattr(model, "Name", "") or "")
    return _LOGICAL_CONTROL_NAMES.get(native_name, native_name)


def _listener_types() -> tuple[type[Any], type[Any]]:
    import unohelper
    from com.sun.star.awt import XActionListener, XKeyHandler

    class ActionListener(unohelper.Base, XActionListener):
        def __init__(self, adapter: InteractiveGameController) -> None:
            self.adapter = adapter

        def actionPerformed(self, event: Any) -> None:  # noqa: N802
            model = getattr(event.Source, "Model", None)
            name = _control_logical_name(model)
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

    return ActionListener, KeyListener


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
    "build_target",
    "create_controller",
]
