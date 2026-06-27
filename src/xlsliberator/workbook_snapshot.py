"""Pure workbook snapshot and diff primitives."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import UTC, datetime
from typing import Any


@dataclass(frozen=True)
class CellSnapshot:
    """Cell state captured from a workbook."""

    sheet: str
    address: str
    value: Any = None
    formula: str | None = None
    error: str | None = None
    type: str | None = None


@dataclass(frozen=True)
class RangeSnapshot:
    """Range state captured from a workbook."""

    sheet: str
    range_address: str
    cells: list[CellSnapshot] = field(default_factory=list)


@dataclass(frozen=True)
class WorkbookSnapshot:
    """Workbook state snapshot."""

    target: str
    backend: str
    ranges: list[RangeSnapshot] = field(default_factory=list)
    timestamp: datetime = field(default_factory=lambda: datetime.now(UTC))
    metadata: dict[str, Any] = field(default_factory=dict)


@dataclass(frozen=True)
class SnapshotDiff:
    """Diff result between two snapshots."""

    matching: int
    mismatching: int
    missing: int
    extra: int
    details: list[dict[str, Any]] = field(default_factory=list)


def compare_cell_values(a: Any, b: Any, tolerance: float = 1e-9) -> bool:
    """Compare cell values with numeric tolerance and empty normalization."""
    if a in (None, "") and b in (None, ""):
        return True
    # bool subclasses int, but a cell that changed from numeric 1 to boolean TRUE
    # (or 0 to FALSE) is a real change. Only treat booleans as equal to booleans.
    if isinstance(a, bool) or isinstance(b, bool):
        return isinstance(a, bool) and isinstance(b, bool) and a == b
    if isinstance(a, int | float) and isinstance(b, int | float):
        return abs(float(a) - float(b)) <= tolerance
    return bool(a == b)


def diff_snapshots(
    before: WorkbookSnapshot,
    after: WorkbookSnapshot,
    expected_changes: list[str] | None = None,
) -> SnapshotDiff:
    """Diff two workbook snapshots."""
    expected = set(expected_changes or [])
    before_cells = _cell_map(before)
    after_cells = _cell_map(after)
    details: list[dict[str, Any]] = []
    matching = 0
    mismatching = 0
    missing = 0
    extra = 0

    for key, before_cell in before_cells.items():
        after_cell = after_cells.get(key)
        if after_cell is None:
            missing += 1
            details.append({"kind": "missing", "cell": key})
            continue
        changed = not compare_cell_values(before_cell.value, after_cell.value)
        if changed and key not in expected:
            mismatching += 1
            details.append(
                {
                    "kind": "mismatch",
                    "cell": key,
                    "before": before_cell.value,
                    "after": after_cell.value,
                }
            )
        else:
            matching += 1

    for key, after_cell in after_cells.items():
        if key in before_cells:
            continue
        extra += 1
        details.append({"kind": "extra", "cell": key, "after": after_cell.value})

    return SnapshotDiff(
        matching=matching,
        mismatching=mismatching,
        missing=missing,
        extra=extra,
        details=details,
    )


def _cell_map(snapshot: WorkbookSnapshot) -> dict[str, CellSnapshot]:
    return {
        f"{cell.sheet}!{cell.address}": cell
        for range_snapshot in snapshot.ranges
        for cell in range_snapshot.cells
    }
