"""Real dataset test: Tippspiel formula equivalence validation."""

from pathlib import Path

import pytest

from xlsliberator.testing_lo import compare_excel_calc


@pytest.mark.real
def test_tippspiel_formula_equivalence() -> None:
    """Test that converted Tippspiel formulas produce equivalent values."""
    # Paths to test files
    test_data_dir = Path(__file__).parent.parent / "data"
    excel_file = test_data_dir / "Bundesliga-Ergebnis-Tippspiel_V2.31_2025-26.xlsm"
    ods_file = Path("output.ods")

    # Skip if output file doesn't exist
    if not ods_file.exists():
        pytest.skip("Output ODS file not found - run conversion first")

    # Compare formulas
    result = compare_excel_calc(excel_file, ods_file, tolerance=1e-9)

    # Assert quality gates
    assert result.formula_cells > 0, "No formula cells found"
    assert result.match_rate >= 95.0, (
        f"Formula match rate {result.match_rate:.2f}% is below 95% threshold. "
        f"Matching: {result.matching}, Mismatching: {result.mismatching}"
    )

    print(result.summary())
