"""WorksheetFunction compatibility adapter."""

from __future__ import annotations

from collections.abc import Iterable


class WorksheetFunctionAdapter:
    """Small pure subset of Excel WorksheetFunction."""

    def sum(self, values: Iterable[float]) -> float:
        """Return the numeric sum of values."""
        return float(sum(values))

    def average(self, values: Iterable[float]) -> float:
        """Return the numeric average of values."""
        collected = list(values)
        if not collected:
            raise ValueError("average requires at least one value")
        return float(sum(collected) / len(collected))
