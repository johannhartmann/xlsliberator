"""Regression tests for non-collapsing formula result comparisons."""

import pytest

from xlsliberator.testing_lo import values_equal


@pytest.mark.parametrize(
    ("source", "target"),
    [
        (None, 0.0),
        (0.0, None),
        (None, ""),
        ("", None),
        (True, 1),
        (False, 0),
        (" text", "text"),
        ("text ", "text"),
        ("#REF!", "#VALUE!"),
    ],
)
def test_semantically_distinct_formula_results_never_compare_equal(
    source: object,
    target: object,
) -> None:
    assert not values_equal(source, target)


def test_numeric_tolerance_is_limited_to_numeric_values() -> None:
    assert values_equal(1.0, 1.0001, tolerance=0.001)
    assert not values_equal(1.0, 1.1, tolerance=0.001)
