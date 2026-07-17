"""Deterministic delta minimizer for synthetic workbook regressions."""

from __future__ import annotations

from collections.abc import Callable
from typing import Literal

from pydantic import BaseModel, ConfigDict, Field


class WorkbookCandidate(BaseModel):
    """Serializable workbook surface understood by the generic minimizer."""

    model_config = ConfigDict(extra="forbid")

    sheets: list[str] = Field(default_factory=list)
    ranges: list[str] = Field(default_factory=list)
    formulas: list[str] = Field(default_factory=list)
    vba_modules: list[str] = Field(default_factory=list)
    vba_procedures: list[str] = Field(default_factory=list)
    package_parts: list[str] = Field(default_factory=list)

    @property
    def size(self) -> int:
        """Count independently removable features."""
        return sum(len(getattr(self, field)) for field in MINIMIZABLE_FIELDS)


MINIMIZABLE_FIELDS = (
    "sheets",
    "ranges",
    "formulas",
    "vba_modules",
    "vba_procedures",
    "package_parts",
)


class MinimizationStep(BaseModel):
    """Auditable accept/reject decision."""

    model_config = ConfigDict(extra="forbid")

    field: str
    removed: list[str]
    retained_failure: bool
    candidate_size: int


class MinimizationEvidence(BaseModel):
    """Evidence that the exact failure signature survived reduction."""

    model_config = ConfigDict(extra="forbid")

    schema_version: Literal["1.0.0"] = "1.0.0"
    failure_signature: str
    original: WorkbookCandidate
    minimized: WorkbookCandidate
    steps: list[MinimizationStep]


FailurePredicate = Callable[[WorkbookCandidate], str | None]


def minimize_failure(
    original: WorkbookCandidate,
    *,
    expected_signature: str,
    predicate: FailurePredicate,
) -> MinimizationEvidence:
    """Remove features one-by-one while preserving the exact trace-diff signature."""
    if predicate(original) != expected_signature:
        raise ValueError("original candidate does not reproduce the expected failure signature")
    current = original.model_copy(deep=True)
    steps: list[MinimizationStep] = []
    for field in MINIMIZABLE_FIELDS:
        index = 0
        while index < len(getattr(current, field)):
            values = list(getattr(current, field))
            removed = [values.pop(index)]
            candidate = current.model_copy(update={field: values}, deep=True)
            retained = predicate(candidate) == expected_signature
            steps.append(
                MinimizationStep(
                    field=field,
                    removed=removed,
                    retained_failure=retained,
                    candidate_size=candidate.size,
                )
            )
            if retained:
                current = candidate
            else:
                index += 1
    if current.size >= original.size:
        raise ValueError("failure could not be minimized")
    if predicate(current) != expected_signature:
        raise RuntimeError("minimizer lost the required failure signature")
    return MinimizationEvidence(
        failure_signature=expected_signature,
        original=original,
        minimized=current,
        steps=steps,
    )
