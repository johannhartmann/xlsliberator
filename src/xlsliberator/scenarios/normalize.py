"""Lossless normalization for source and target observations."""

from __future__ import annotations

from datetime import date, datetime
from typing import Any, Literal

from xlsliberator.scenarios.models import ObservationValue, ValueKind

KNOWN_ERRORS = {
    "#BLOCKED!",
    "#CALC!",
    "#BUSY!",
    "#CONNECT!",
    "#DATA!",
    "#DIV/0!",
    "#FIELD!",
    "#GETTING_DATA",
    "#N/A",
    "#NAME?",
    "#NULL!",
    "#NUM!",
    "#PYTHON!",
    "#REF!",
    "#SPILL!",
    "#UNKNOWN!",
    "#VALUE!",
}


def normalize_value(
    value: Any,
    *,
    date_system: Literal["1900", "1904"] = "1900",
    timezone: str = "UTC",
    formula: str | None = None,
    cell_type: str | None = None,
    error_type: str | None = None,
) -> ObservationValue:
    """Normalize a runtime value without collapsing semantically distinct cases."""
    common: dict[str, Any] = {"formula": formula, "cell_type": cell_type}
    if error_type is not None:
        return ObservationValue(
            kind=ValueKind.ERROR,
            value=value,
            error_type=error_type,
            **common,
        )
    if value is None:
        return ObservationValue(kind=ValueKind.EMPTY_CELL, value=None, **common)
    if isinstance(value, bool):
        return ObservationValue(kind=ValueKind.BOOLEAN, value=value, **common)
    if isinstance(value, str):
        if value in KNOWN_ERRORS:
            return ObservationValue(kind=ValueKind.ERROR, value=value, error_type=value, **common)
        if value == "":
            return ObservationValue(kind=ValueKind.EMPTY_STRING, value="", **common)
        return ObservationValue(kind=ValueKind.STRING, value=value, **common)
    if isinstance(value, datetime):
        return ObservationValue(
            kind=ValueKind.DATETIME,
            value=value.isoformat(),
            date_system=date_system,
            timezone=timezone,
            **common,
        )
    if isinstance(value, date):
        return ObservationValue(
            kind=ValueKind.DATE,
            value=value.isoformat(),
            date_system=date_system,
            timezone=timezone,
            **common,
        )
    if isinstance(value, (int, float)):
        return ObservationValue(kind=ValueKind.NUMBER, value=value, **common)
    if isinstance(value, (list, tuple)):
        return ObservationValue(kind=ValueKind.ARRAY, value=list(value), **common)
    if isinstance(value, dict):
        return ObservationValue(kind=ValueKind.OBJECT, value=value, **common)
    return ObservationValue(
        kind=ValueKind.OBJECT,
        value=repr(value),
        metadata={"python_type": type(value).__qualname__},
        **common,
    )
