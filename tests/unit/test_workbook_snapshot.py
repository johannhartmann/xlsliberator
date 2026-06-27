"""Tests for pure workbook snapshot diff primitives."""

from xlsliberator.workbook_snapshot import (
    CellSnapshot,
    RangeSnapshot,
    WorkbookSnapshot,
    compare_cell_values,
    diff_snapshots,
)


def test_compare_cell_values_numeric_tolerance() -> None:
    """Numeric values should compare with tolerance."""
    assert compare_cell_values(1.0, 1.0 + 1e-10)
    assert not compare_cell_values(1.0, 1.1)


def test_compare_cell_values_empty_normalization() -> None:
    """None and empty strings should compare as empty."""
    assert compare_cell_values(None, "")
    assert not compare_cell_values(None, "value")


def test_diff_snapshots_reports_mismatch_missing_and_extra() -> None:
    """Snapshot diffs should include mismatch details."""
    before = WorkbookSnapshot(
        target="before.ods",
        backend="test",
        ranges=[
            RangeSnapshot(
                sheet="Sheet1",
                range_address="A1:B1",
                cells=[
                    CellSnapshot(sheet="Sheet1", address="A1", value=1),
                    CellSnapshot(sheet="Sheet1", address="B1", value=2),
                ],
            )
        ],
    )
    after = WorkbookSnapshot(
        target="after.ods",
        backend="test",
        ranges=[
            RangeSnapshot(
                sheet="Sheet1",
                range_address="A1:C1",
                cells=[
                    CellSnapshot(sheet="Sheet1", address="A1", value=1),
                    CellSnapshot(sheet="Sheet1", address="B1", value=3),
                    CellSnapshot(sheet="Sheet1", address="C1", value=4),
                ],
            )
        ],
    )

    diff = diff_snapshots(before, after)

    assert diff.matching == 1
    assert diff.mismatching == 1
    assert diff.extra == 1
    assert any(detail["kind"] == "mismatch" for detail in diff.details)


def test_diff_snapshots_expected_change_counts_as_matching() -> None:
    """Expected changes should not be reported as mismatches."""
    before = WorkbookSnapshot(
        target="before.ods",
        backend="test",
        ranges=[RangeSnapshot("Sheet1", "A1", [CellSnapshot("Sheet1", "A1", value=1)])],
    )
    after = WorkbookSnapshot(
        target="after.ods",
        backend="test",
        ranges=[RangeSnapshot("Sheet1", "A1", [CellSnapshot("Sheet1", "A1", value=2)])],
    )

    diff = diff_snapshots(before, after, expected_changes=["Sheet1!A1"])

    assert diff.matching == 1
    assert diff.mismatching == 0


def test_compare_cell_values_bool_is_not_numeric() -> None:
    """A numeric 1/0 changing to boolean TRUE/FALSE must count as a change."""
    assert not compare_cell_values(1, True)
    assert not compare_cell_values(0, False)
    assert not compare_cell_values(True, 1)
    # Booleans still compare equal to booleans.
    assert compare_cell_values(True, True)
    assert not compare_cell_values(True, False)


def test_diff_snapshots_flags_numeric_to_bool_change() -> None:
    """A cell flipping from numeric 1 to boolean TRUE is a real mismatch."""
    before = WorkbookSnapshot(
        target="before.ods",
        backend="test",
        ranges=[RangeSnapshot("Sheet1", "A1", [CellSnapshot("Sheet1", "A1", value=1)])],
    )
    after = WorkbookSnapshot(
        target="after.ods",
        backend="test",
        ranges=[RangeSnapshot("Sheet1", "A1", [CellSnapshot("Sheet1", "A1", value=True)])],
    )

    diff = diff_snapshots(before, after)

    assert diff.mismatching == 1
    assert diff.matching == 0
