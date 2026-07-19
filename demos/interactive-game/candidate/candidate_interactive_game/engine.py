"""Deterministic, serializable game engine for the interactive-game migration.

The engine is deliberately independent from LibreOffice and UNO.  A target
adapter may translate keyboard, timer, control, and document events into these
pure state transitions and render the resulting cells in Calc.
"""

from __future__ import annotations

import json
from collections.abc import Mapping
from dataclasses import dataclass, replace
from enum import StrEnum
from typing import Final, cast

SCHEMA_VERSION: Final = "1.0.0"
BOARD_ROWS: Final = 28
BOARD_COLUMNS: Final = 15
SPAWN_ROW: Final = 0
SPAWN_COLUMN: Final = 6

SOFT_DROP_POINTS: Final = 1
LANDING_POINTS: Final = 5
LINE_POINTS: Final = 100

_UINT64_MASK: Final = (1 << 64) - 1
_LCG_MULTIPLIER: Final = 6_364_136_223_846_793_005
_LCG_INCREMENT: Final = 1_442_695_040_888_963_407


class GamePhase(StrEnum):
    """User-visible lifecycle of one game."""

    STOPPED = "stopped"
    RUNNING = "running"
    PAUSED = "paused"
    GAME_OVER = "game_over"


class PieceKind(StrEnum):
    """The nine source-workbook piece types."""

    L_LEFT = "l_left"
    L_RIGHT = "l_right"
    LINE = "line"
    T = "t"
    SQUARE = "square"
    DOMINO = "domino"
    SINGLE = "single"
    TWIST_LEFT = "twist_left"
    TWIST_RIGHT = "twist_right"


Coordinate = tuple[int, int]
Shape = tuple[Coordinate, ...]

_BASE_SHAPES: Final[dict[PieceKind, Shape]] = {
    PieceKind.L_LEFT: ((0, 0), (1, 0), (2, 0), (2, 1)),
    PieceKind.L_RIGHT: ((0, 0), (1, 0), (2, 0), (2, -1)),
    PieceKind.LINE: ((0, 0), (1, 0), (2, 0), (3, 0)),
    PieceKind.T: ((0, 0), (1, 0), (2, 0), (1, 1)),
    PieceKind.SQUARE: ((0, 0), (1, 0), (0, 1), (1, 1)),
    PieceKind.DOMINO: ((0, 0), (1, 0)),
    PieceKind.SINGLE: ((0, 0),),
    PieceKind.TWIST_LEFT: ((0, 0), (1, 0), (1, 1), (2, 1)),
    PieceKind.TWIST_RIGHT: ((0, 1), (1, 1), (1, 0), (2, 0)),
}

_PIECE_COLORS: Final[dict[PieceKind, int]] = {
    PieceKind.L_LEFT: 3,
    PieceKind.L_RIGHT: 3,
    PieceKind.LINE: 6,
    PieceKind.T: 4,
    PieceKind.SQUARE: 41,
    PieceKind.DOMINO: 48,
    PieceKind.SINGLE: 45,
    PieceKind.TWIST_LEFT: 54,
    PieceKind.TWIST_RIGHT: 54,
}

_PIECE_ORDER: Final[tuple[PieceKind, ...]] = tuple(PieceKind)


def _normalize_shape(shape: Shape) -> Shape:
    minimum_row = min(row for row, _column in shape)
    minimum_column = min(column for _row, column in shape)
    return tuple(sorted((row - minimum_row, column - minimum_column) for row, column in shape))


def _rotate_shape(shape: Shape) -> Shape:
    return _normalize_shape(tuple((column, -row) for row, column in shape))


def _orientations(shape: Shape) -> tuple[Shape, ...]:
    orientations: list[Shape] = []
    current = _normalize_shape(shape)
    while current not in orientations:
        orientations.append(current)
        current = _rotate_shape(current)
    return tuple(orientations)


_SHAPES: Final = {kind: _orientations(shape) for kind, shape in _BASE_SHAPES.items()}


@dataclass(frozen=True, order=True, slots=True)
class SettledCell:
    """One occupied board position."""

    row: int
    column: int
    color: int

    def __post_init__(self) -> None:
        if not 0 <= self.row < BOARD_ROWS:
            raise ValueError(f"settled cell row is outside the board: {self.row}")
        if not 0 <= self.column < BOARD_COLUMNS:
            raise ValueError(f"settled cell column is outside the board: {self.column}")
        if self.color <= 0:
            raise ValueError("settled cell color must be positive")


@dataclass(frozen=True, slots=True)
class ActivePiece:
    """Current falling piece, anchored at the top-left of its orientation."""

    kind: PieceKind
    rotation: int
    row: int
    column: int

    def __post_init__(self) -> None:
        if not 0 <= self.rotation < len(_SHAPES[self.kind]):
            raise ValueError(f"invalid rotation {self.rotation} for {self.kind.value}")


@dataclass(frozen=True, slots=True)
class GameState:
    """Complete state required to continue a game after save and reopen."""

    phase: GamePhase
    settled: tuple[SettledCell, ...]
    active: ActivePiece | None
    next_piece: PieceKind | None
    score: int
    completed_lines: int
    high_score: int
    tick_index: int
    event_index: int
    seed: int
    rng_state: int
    draw_index: int

    def __post_init__(self) -> None:
        for name, value in (
            ("score", self.score),
            ("completed_lines", self.completed_lines),
            ("high_score", self.high_score),
            ("tick_index", self.tick_index),
            ("event_index", self.event_index),
            ("draw_index", self.draw_index),
        ):
            if value < 0:
                raise ValueError(f"{name} must not be negative")
        for name, value in (("seed", self.seed), ("rng_state", self.rng_state)):
            if not 0 <= value <= _UINT64_MASK:
                raise ValueError(f"{name} must fit in an unsigned 64-bit integer")

        positions = [(cell.row, cell.column) for cell in self.settled]
        if len(positions) != len(set(positions)):
            raise ValueError("settled cells contain duplicate positions")
        if tuple(sorted(self.settled)) != self.settled:
            raise ValueError("settled cells must use deterministic sorted order")

        if self.phase in {GamePhase.RUNNING, GamePhase.PAUSED}:
            if self.active is None or self.next_piece is None:
                raise ValueError("an active game requires active and next pieces")
            if not _piece_fits(self.active, self.settled):
                raise ValueError("active piece is outside the board or collides with settled cells")
        elif self.active is not None:
            raise ValueError("stopped and game-over states cannot have an active piece")


def new_game_state(*, seed: int = 1, high_score: int = 0) -> GameState:
    """Create a stopped, empty game with a reproducible random stream."""
    if not 0 <= seed <= _UINT64_MASK:
        raise ValueError("seed must fit in an unsigned 64-bit integer")
    if high_score < 0:
        raise ValueError("high_score must not be negative")
    return GameState(
        phase=GamePhase.STOPPED,
        settled=(),
        active=None,
        next_piece=None,
        score=0,
        completed_lines=0,
        high_score=high_score,
        tick_index=0,
        event_index=0,
        seed=seed,
        rng_state=seed,
        draw_index=0,
    )


def start_game(state: GameState) -> GameState:
    """Start a fresh game, preserving only the high score and configured seed."""
    if state.phase is GamePhase.RUNNING:
        return state
    base = new_game_state(seed=state.seed, high_score=state.high_score)
    first, base = _draw_piece(base)
    following, base = _draw_piece(base)
    return replace(
        base,
        phase=GamePhase.RUNNING,
        active=_spawn_piece(first),
        next_piece=following,
        event_index=state.event_index + 1,
    )


def pause_game(state: GameState) -> GameState:
    """Pause timer-driven progress without changing board state."""
    if state.phase is not GamePhase.RUNNING:
        return state
    return replace(state, phase=GamePhase.PAUSED, event_index=state.event_index + 1)


def resume_game(state: GameState) -> GameState:
    """Resume a paused game."""
    if state.phase is not GamePhase.PAUSED:
        return state
    return replace(state, phase=GamePhase.RUNNING, event_index=state.event_index + 1)


def toggle_pause(state: GameState) -> GameState:
    """Apply the migrated Pause/Resume control behavior."""
    if state.phase is GamePhase.RUNNING:
        return pause_game(state)
    if state.phase is GamePhase.PAUSED:
        return resume_game(state)
    return state


def reset_game(state: GameState) -> GameState:
    """Clear board, score, and progress while retaining the high score."""
    reset = new_game_state(seed=state.seed, high_score=state.high_score)
    return replace(reset, event_index=state.event_index + 1)


def move_left(state: GameState) -> GameState:
    """Move the active piece one column left when valid."""
    return _move(state, row_delta=0, column_delta=-1)


def move_right(state: GameState) -> GameState:
    """Move the active piece one column right when valid."""
    return _move(state, row_delta=0, column_delta=1)


def soft_drop(state: GameState) -> GameState:
    """Apply the source workbook's two-row quick drop and one bonus point."""
    if state.phase is not GamePhase.RUNNING or state.active is None:
        return state
    candidate = replace(state.active, row=state.active.row + 2)
    if not _piece_fits(candidate, state.settled):
        return state
    score = state.score + SOFT_DROP_POINTS
    return replace(
        state,
        active=candidate,
        score=score,
        high_score=max(state.high_score, score),
        event_index=state.event_index + 1,
    )


def rotate_piece(state: GameState) -> GameState:
    """Rotate clockwise once, rejecting rotations that collide or cross a wall."""
    if state.phase is not GamePhase.RUNNING or state.active is None:
        return state
    orientation_count = len(_SHAPES[state.active.kind])
    if orientation_count == 1:
        return state
    candidate = replace(
        state.active,
        rotation=(state.active.rotation + 1) % orientation_count,
    )
    if not _piece_fits(candidate, state.settled):
        return state
    return replace(state, active=candidate, event_index=state.event_index + 1)


def tick(state: GameState) -> GameState:
    """Perform at most one bounded falling-piece update."""
    if state.phase is not GamePhase.RUNNING or state.active is None:
        return state

    candidate = replace(state.active, row=state.active.row + 1)
    if _piece_fits(candidate, state.settled):
        return replace(
            state,
            active=candidate,
            tick_index=state.tick_index + 1,
            event_index=state.event_index + 1,
        )
    return _land_and_spawn(state)


def tick_interval_ms(state: GameState) -> int:
    """Return the source-derived timer interval for completed-line progress."""
    lines = state.completed_lines
    if lines <= 9:
        return 160
    if lines <= 19:
        return 145
    if lines <= 29:
        return 120
    if lines <= 39:
        return 110
    if lines <= 49:
        return 100
    if lines <= 59:
        return 90
    if lines <= 69:
        return 70
    if lines <= 79:
        return 60
    if lines <= 99:
        return 50
    return 40


def piece_cells(piece: ActivePiece) -> tuple[Coordinate, ...]:
    """Return absolute board positions occupied by an active piece."""
    return tuple(
        (piece.row + row, piece.column + column)
        for row, column in _SHAPES[piece.kind][piece.rotation]
    )


def state_to_dict(state: GameState) -> dict[str, object]:
    """Serialize state into a stable JSON-compatible mapping."""
    active: dict[str, object] | None = None
    if state.active is not None:
        active = {
            "kind": state.active.kind.value,
            "rotation": state.active.rotation,
            "row": state.active.row,
            "column": state.active.column,
        }
    return {
        "schema_version": SCHEMA_VERSION,
        "phase": state.phase.value,
        "settled": [
            {"row": cell.row, "column": cell.column, "color": cell.color} for cell in state.settled
        ],
        "active": active,
        "next_piece": state.next_piece.value if state.next_piece is not None else None,
        "score": state.score,
        "completed_lines": state.completed_lines,
        "high_score": state.high_score,
        "tick_index": state.tick_index,
        "event_index": state.event_index,
        "seed": state.seed,
        "rng_state": state.rng_state,
        "draw_index": state.draw_index,
    }


def state_from_dict(payload: Mapping[str, object]) -> GameState:
    """Deserialize and validate a mapping produced by :func:`state_to_dict`."""
    expected_keys = {
        "schema_version",
        "phase",
        "settled",
        "active",
        "next_piece",
        "score",
        "completed_lines",
        "high_score",
        "tick_index",
        "event_index",
        "seed",
        "rng_state",
        "draw_index",
    }
    if set(payload) != expected_keys:
        missing = sorted(expected_keys - set(payload))
        unknown = sorted(set(payload) - expected_keys)
        raise ValueError(f"invalid state fields; missing={missing}, unknown={unknown}")
    if payload["schema_version"] != SCHEMA_VERSION:
        raise ValueError(f"unsupported game-state schema: {payload['schema_version']!r}")

    settled_payload = payload["settled"]
    if not isinstance(settled_payload, list):
        raise ValueError("settled must be a list")
    settled = tuple(sorted(_settled_cell_from_payload(item) for item in settled_payload))

    active_payload = payload["active"]
    active = None
    if active_payload is not None:
        active_mapping = _as_mapping(active_payload, "active")
        if set(active_mapping) != {"kind", "rotation", "row", "column"}:
            raise ValueError("active has invalid fields")
        active = ActivePiece(
            kind=PieceKind(_mapping_string(active_mapping, "kind")),
            rotation=_mapping_int(active_mapping, "rotation"),
            row=_mapping_int(active_mapping, "row"),
            column=_mapping_int(active_mapping, "column"),
        )

    next_piece_payload = payload["next_piece"]
    if next_piece_payload is not None and not isinstance(next_piece_payload, str):
        raise ValueError("next_piece must be a string or null")
    return GameState(
        phase=GamePhase(_mapping_string(payload, "phase")),
        settled=settled,
        active=active,
        next_piece=PieceKind(next_piece_payload) if next_piece_payload is not None else None,
        score=_mapping_int(payload, "score"),
        completed_lines=_mapping_int(payload, "completed_lines"),
        high_score=_mapping_int(payload, "high_score"),
        tick_index=_mapping_int(payload, "tick_index"),
        event_index=_mapping_int(payload, "event_index"),
        seed=_mapping_int(payload, "seed"),
        rng_state=_mapping_int(payload, "rng_state"),
        draw_index=_mapping_int(payload, "draw_index"),
    )


def state_to_json(state: GameState) -> str:
    """Serialize state deterministically for the hidden ODS state sheet."""
    return json.dumps(state_to_dict(state), sort_keys=True, separators=(",", ":"))


def state_from_json(payload: str) -> GameState:
    """Load and validate deterministic state JSON."""
    decoded = cast(object, json.loads(payload))
    if not isinstance(decoded, Mapping):
        raise ValueError("game-state JSON must contain an object")
    if not all(isinstance(key, str) for key in decoded):
        raise ValueError("game-state JSON keys must be strings")
    return state_from_dict(cast(Mapping[str, object], decoded))


def _move(state: GameState, *, row_delta: int, column_delta: int) -> GameState:
    if state.phase is not GamePhase.RUNNING or state.active is None:
        return state
    candidate = replace(
        state.active,
        row=state.active.row + row_delta,
        column=state.active.column + column_delta,
    )
    if not _piece_fits(candidate, state.settled):
        return state
    return replace(state, active=candidate, event_index=state.event_index + 1)


def _land_and_spawn(state: GameState) -> GameState:
    active = state.active
    if active is None or state.next_piece is None:
        raise ValueError("cannot land an incomplete running state")
    color = _PIECE_COLORS[active.kind]
    occupied = {(cell.row, cell.column): cell for cell in state.settled}
    for row, column in piece_cells(active):
        occupied[(row, column)] = SettledCell(row=row, column=column, color=color)
    settled = tuple(sorted(occupied.values()))

    settled, removed_lines = _collapse_complete_lines(settled)
    score = state.score + LANDING_POINTS + removed_lines * LINE_POINTS
    completed_lines = state.completed_lines + removed_lines
    following, advanced = _draw_piece(state)
    spawned = _spawn_piece(state.next_piece)
    if not _piece_fits(spawned, settled):
        return replace(
            state,
            phase=GamePhase.GAME_OVER,
            settled=settled,
            active=None,
            next_piece=following,
            score=score,
            completed_lines=completed_lines,
            high_score=max(state.high_score, score),
            tick_index=state.tick_index + 1,
            event_index=state.event_index + 1,
            rng_state=advanced.rng_state,
            draw_index=advanced.draw_index,
        )
    return replace(
        state,
        settled=settled,
        active=spawned,
        next_piece=following,
        score=score,
        completed_lines=completed_lines,
        high_score=max(state.high_score, score),
        tick_index=state.tick_index + 1,
        event_index=state.event_index + 1,
        rng_state=advanced.rng_state,
        draw_index=advanced.draw_index,
    )


def _collapse_complete_lines(
    settled: tuple[SettledCell, ...],
) -> tuple[tuple[SettledCell, ...], int]:
    columns_by_row: dict[int, set[int]] = {}
    for cell in settled:
        columns_by_row.setdefault(cell.row, set()).add(cell.column)
    complete_rows = {
        row for row, columns in columns_by_row.items() if len(columns) == BOARD_COLUMNS
    }
    if not complete_rows:
        return settled, 0

    collapsed: list[SettledCell] = []
    for cell in settled:
        if cell.row in complete_rows:
            continue
        rows_removed_below = sum(complete_row > cell.row for complete_row in complete_rows)
        collapsed.append(replace(cell, row=cell.row + rows_removed_below))
    return tuple(sorted(collapsed)), len(complete_rows)


def _piece_fits(piece: ActivePiece, settled: tuple[SettledCell, ...]) -> bool:
    positions = piece_cells(piece)
    if any(
        row < 0 or row >= BOARD_ROWS or column < 0 or column >= BOARD_COLUMNS
        for row, column in positions
    ):
        return False
    settled_positions = {(cell.row, cell.column) for cell in settled}
    return not any(position in settled_positions for position in positions)


def _spawn_piece(kind: PieceKind) -> ActivePiece:
    return ActivePiece(
        kind=kind,
        rotation=0,
        row=SPAWN_ROW,
        column=SPAWN_COLUMN,
    )


def _draw_piece(state: GameState) -> tuple[PieceKind, GameState]:
    rng_state = (state.rng_state * _LCG_MULTIPLIER + _LCG_INCREMENT) & _UINT64_MASK
    kind = _PIECE_ORDER[rng_state % len(_PIECE_ORDER)]
    return kind, replace(
        state,
        rng_state=rng_state,
        draw_index=state.draw_index + 1,
    )


def _as_mapping(value: object, name: str) -> Mapping[str, object]:
    if not isinstance(value, Mapping):
        raise ValueError(f"{name} must be an object")
    if not all(isinstance(key, str) for key in value):
        raise ValueError(f"{name} keys must be strings")
    return cast(Mapping[str, object], value)


def _settled_cell_from_payload(value: object) -> SettledCell:
    mapping = _as_mapping(value, "settled cell")
    if set(mapping) != {"row", "column", "color"}:
        raise ValueError("settled cell has invalid fields")
    return SettledCell(
        row=_mapping_int(mapping, "row"),
        column=_mapping_int(mapping, "column"),
        color=_mapping_int(mapping, "color"),
    )


def _mapping_int(mapping: Mapping[str, object], key: str) -> int:
    value = mapping.get(key)
    if isinstance(value, bool) or not isinstance(value, int):
        raise ValueError(f"{key} must be an integer")
    return value


def _mapping_string(mapping: Mapping[str, object], key: str) -> str:
    value = mapping.get(key)
    if not isinstance(value, str):
        raise ValueError(f"{key} must be a string")
    return value


__all__ = [
    "BOARD_COLUMNS",
    "BOARD_ROWS",
    "LANDING_POINTS",
    "LINE_POINTS",
    "SCHEMA_VERSION",
    "SOFT_DROP_POINTS",
    "ActivePiece",
    "GamePhase",
    "GameState",
    "PieceKind",
    "SettledCell",
    "move_left",
    "move_right",
    "new_game_state",
    "pause_game",
    "piece_cells",
    "reset_game",
    "resume_game",
    "rotate_piece",
    "soft_drop",
    "start_game",
    "state_from_dict",
    "state_from_json",
    "state_to_dict",
    "state_to_json",
    "tick",
    "tick_interval_ms",
    "toggle_pause",
]
