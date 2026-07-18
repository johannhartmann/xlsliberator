"""Source-derived behavioral tests for the interactive game engine."""

from dataclasses import replace

import pytest

from xlsliberator.interactive_game_engine import (
    BOARD_COLUMNS,
    LANDING_POINTS,
    LINE_POINTS,
    SOFT_DROP_POINTS,
    ActivePiece,
    GamePhase,
    GameState,
    PieceKind,
    SettledCell,
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
    state_from_dict,
    state_from_json,
    state_to_dict,
    state_to_json,
    tick,
    tick_interval_ms,
    toggle_pause,
)


def _running_single(
    *,
    row: int,
    column: int,
    settled: tuple[SettledCell, ...] = (),
    score: int = 0,
    completed_lines: int = 0,
    high_score: int = 0,
) -> GameState:
    state = start_game(new_game_state(seed=17, high_score=high_score))
    return replace(
        state,
        active=ActivePiece(PieceKind.SINGLE, rotation=0, row=row, column=column),
        next_piece=PieceKind.SINGLE,
        settled=tuple(sorted(settled)),
        score=score,
        completed_lines=completed_lines,
    )


def test_seeded_start_is_deterministic_and_serializable() -> None:
    first = start_game(new_game_state(seed=1234, high_score=800))
    second = start_game(new_game_state(seed=1234, high_score=800))

    assert first == second
    assert first.phase is GamePhase.RUNNING
    assert first.active is not None
    assert first.next_piece is not None
    assert first.draw_index == 2

    encoded = state_to_json(first)
    assert encoded == state_to_json(second)
    assert state_from_json(encoded) == first
    assert state_from_dict(state_to_dict(first)) == first


def test_deserialization_rejects_unknown_or_malformed_state() -> None:
    payload = state_to_dict(new_game_state())
    payload["unexpected"] = True

    with pytest.raises(ValueError, match="invalid state fields"):
        state_from_dict(payload)
    with pytest.raises(ValueError, match="must contain an object"):
        state_from_json("[]")

    malformed_cell = state_to_dict(new_game_state())
    malformed_cell["settled"] = [{"row": 1, "column": 2, "color": 3, "extra": 4}]
    with pytest.raises(ValueError, match="settled cell has invalid fields"):
        state_from_dict(malformed_cell)


def test_start_pause_resume_toggle_and_reset_transitions() -> None:
    stopped = new_game_state(seed=9, high_score=250)
    running = start_game(stopped)
    paused = pause_game(running)
    resumed = resume_game(paused)
    toggled = toggle_pause(resumed)
    reset = reset_game(replace(toggled, score=125, high_score=250))

    assert paused.phase is GamePhase.PAUSED
    assert tick(paused) is paused
    assert resumed.phase is GamePhase.RUNNING
    assert toggled.phase is GamePhase.PAUSED
    assert reset.phase is GamePhase.STOPPED
    assert reset.active is None
    assert reset.settled == ()
    assert reset.score == 0
    assert reset.completed_lines == 0
    assert reset.high_score == 250
    assert reset.rng_state == reset.seed
    assert reset.draw_index == 0


def test_horizontal_movement_changes_state_once_and_wall_collision_is_noop() -> None:
    state = _running_single(row=3, column=0)
    blocked_by_wall = move_left(state)
    moved = move_right(state)
    occupied = replace(
        state,
        settled=(SettledCell(row=3, column=1, color=3),),
    )

    assert blocked_by_wall is state
    assert moved.active == ActivePiece(PieceKind.SINGLE, rotation=0, row=3, column=1)
    assert moved.event_index == state.event_index + 1
    assert move_right(occupied) is occupied


def test_rotation_is_clockwise_and_rejected_at_wall_or_settled_cell() -> None:
    started = start_game(new_game_state(seed=2))
    vertical = replace(
        started,
        active=ActivePiece(PieceKind.LINE, rotation=0, row=3, column=6),
    )
    rotated = rotate_piece(vertical)
    wall_blocked = replace(
        vertical,
        active=ActivePiece(PieceKind.LINE, rotation=0, row=3, column=14),
    )
    cell_blocked = replace(
        vertical,
        settled=(SettledCell(row=3, column=7, color=41),),
    )

    assert rotated.active == ActivePiece(PieceKind.LINE, rotation=1, row=3, column=6)
    assert piece_cells(rotated.active) == ((3, 6), (3, 7), (3, 8), (3, 9))
    assert rotate_piece(wall_blocked) is wall_blocked
    assert rotate_piece(cell_blocked) is cell_blocked


def test_soft_drop_moves_two_rows_and_awards_only_a_valid_input() -> None:
    state = _running_single(row=4, column=5, score=10, high_score=10)
    dropped = soft_drop(state)
    blocked = replace(
        state,
        settled=(SettledCell(row=6, column=5, color=4),),
    )

    assert dropped.active == ActivePiece(PieceKind.SINGLE, rotation=0, row=6, column=5)
    assert dropped.score == 10 + SOFT_DROP_POINTS
    assert dropped.high_score == 10 + SOFT_DROP_POINTS
    assert dropped.event_index == state.event_index + 1
    assert soft_drop(blocked) is blocked


def test_tick_performs_one_bounded_fall_update() -> None:
    state = _running_single(row=4, column=5)

    advanced = tick(state)

    assert advanced.active == ActivePiece(PieceKind.SINGLE, rotation=0, row=5, column=5)
    assert advanced.tick_index == state.tick_index + 1
    assert advanced.event_index == state.event_index + 1
    assert advanced.score == state.score
    assert advanced.settled == state.settled


def test_landing_locks_piece_scores_and_spawns_declared_next_piece() -> None:
    state = _running_single(
        row=27,
        column=4,
        score=20,
        high_score=20,
    )

    landed = tick(state)

    assert SettledCell(row=27, column=4, color=45) in landed.settled
    assert landed.score == 20 + LANDING_POINTS
    assert landed.high_score == 20 + LANDING_POINTS
    assert landed.active is not None
    assert landed.active.kind is PieceKind.SINGLE
    assert landed.active.row == 0
    assert landed.tick_index == state.tick_index + 1
    assert landed.draw_index == state.draw_index + 1


def test_completed_line_is_removed_and_every_upper_cell_moves_down_once() -> None:
    bottom = tuple(
        SettledCell(row=27, column=column, color=3) for column in range(BOARD_COLUMNS - 1)
    )
    marker = SettledCell(row=25, column=0, color=54)
    state = _running_single(
        row=26,
        column=BOARD_COLUMNS - 1,
        settled=(*bottom, marker),
        score=7,
        completed_lines=3,
        high_score=7,
    )

    positioned = tick(state)
    collapsed = tick(positioned)

    assert positioned.active == ActivePiece(
        PieceKind.SINGLE,
        rotation=0,
        row=27,
        column=BOARD_COLUMNS - 1,
    )
    assert collapsed.score == 7 + LANDING_POINTS + LINE_POINTS
    assert collapsed.completed_lines == 4
    assert SettledCell(row=26, column=0, color=54) in collapsed.settled
    assert not any(cell.row == 27 and cell.color == 3 for cell in collapsed.settled)


def test_two_completed_lines_shift_an_upper_cell_exactly_two_rows() -> None:
    complete_rows = tuple(
        SettledCell(row=row, column=column, color=6)
        for row in (26, 27)
        for column in range(BOARD_COLUMNS)
    )
    state = _running_single(
        row=25,
        column=0,
        settled=complete_rows,
        score=0,
        high_score=0,
    )

    collapsed = tick(state)

    assert collapsed.score == LANDING_POINTS + 2 * LINE_POINTS
    assert collapsed.completed_lines == 2
    assert collapsed.settled == (SettledCell(row=27, column=0, color=45),)


def test_blocked_spawn_finishes_game_and_persists_high_score() -> None:
    state = _running_single(
        row=27,
        column=4,
        settled=(SettledCell(row=0, column=6, color=3),),
        score=95,
        high_score=100,
    )

    finished = tick(state)
    restored = state_from_json(state_to_json(finished))

    assert finished.phase is GamePhase.GAME_OVER
    assert finished.active is None
    assert finished.score == 100
    assert finished.high_score == 100
    assert restored == finished


@pytest.mark.parametrize(
    ("completed_lines", "expected_ms"),
    [
        (0, 160),
        (9, 160),
        (10, 145),
        (20, 120),
        (30, 110),
        (40, 100),
        (50, 90),
        (60, 70),
        (70, 60),
        (80, 50),
        (99, 50),
        (100, 40),
    ],
)
def test_timer_interval_preserves_source_progression(
    completed_lines: int,
    expected_ms: int,
) -> None:
    state = replace(new_game_state(), completed_lines=completed_lines)

    assert tick_interval_ms(state) == expected_ms
