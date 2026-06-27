"""Formula validation seam for source and target formulas."""

from collections.abc import Callable
from enum import StrEnum
from typing import Any, cast

from pydantic import BaseModel, Field

from xlsliberator.calc_backend import CalcBackend
from xlsliberator.validation_models import (
    FormulaIR,
    SourceRef,
    TargetRef,
    ValidationSeverity,
    WorkbookArtifactIR,
)


class FormulaDialect(StrEnum):
    """Supported formula dialect labels."""

    EXCEL_A1 = "excel_a1"
    EXCEL_R1C1 = "excel_r1c1"
    CALC_A1 = "calc_a1"
    OPENFORMULA = "openformula"


class FormulaParseResult(BaseModel):
    """Basic formula parse/validation result."""

    success: bool
    formula: str
    dialect: FormulaDialect
    tokens: list[str] | None = None
    error: str | None = None
    details: dict[str, Any] = Field(default_factory=dict)


class FormulaValidationResult(BaseModel):
    """Formula validation result tied to source/target refs."""

    source_ref: SourceRef
    target_ref: TargetRef | None = None
    parse_result: FormulaParseResult
    recalculation_result: dict[str, Any] | None = None
    severity: ValidationSeverity = ValidationSeverity.INFO


class FormulaEngine:
    """Conservative formula validation interface.

    This class currently performs basic structural checks only. Full target validation
    should use a backend UNO FormulaParser integration.
    """

    def collect_formulas(self, inventory: WorkbookArtifactIR) -> list[FormulaIR]:
        """Collect formula IR entries from a workbook artifact inventory."""
        return list(inventory.formulas)

    def validate_formula_text(
        self,
        formula: str,
        dialect: FormulaDialect,
    ) -> FormulaParseResult:
        """Validate formula text with conservative structural checks."""
        if not formula.startswith("="):
            return FormulaParseResult(
                success=False,
                formula=formula,
                dialect=dialect,
                error="Formula must start with '='",
                details={"validation_scope": "basic_structural"},
            )

        balance_error = _balanced_parentheses_and_quotes_error(formula)
        if balance_error:
            return FormulaParseResult(
                success=False,
                formula=formula,
                dialect=dialect,
                error=balance_error,
                details={"validation_scope": "basic_structural"},
            )

        return FormulaParseResult(
            success=True,
            formula=formula,
            dialect=dialect,
            tokens=None,
            details={
                "validation_scope": "basic_structural",
                "note": "Full target FormulaParser validation is not implemented in this method",
            },
        )

    def validate_target_formula_with_backend(
        self,
        formula: str,
        backend: CalcBackend,
    ) -> FormulaParseResult:
        """Validate target formula with a backend when it implements parsing."""
        parse_formula = getattr(backend, "parse_formula_text", None)
        if callable(parse_formula):
            typed_parse = cast(Callable[[str], FormulaParseResult], parse_formula)
            return typed_parse(formula)

        return FormulaParseResult(
            success=False,
            formula=formula,
            dialect=FormulaDialect.CALC_A1,
            error="Backend FormulaParser integration is not available",
            details={
                "backend_kind": backend.info.kind.value,
                "backend_version": backend.info.version,
            },
        )


def _balanced_parentheses_and_quotes_error(formula: str) -> str | None:
    in_string = False
    depth = 0
    index = 0
    while index < len(formula):
        char = formula[index]
        if char == '"':
            if in_string and index + 1 < len(formula) and formula[index + 1] == '"':
                index += 2
                continue
            in_string = not in_string
        elif not in_string:
            if char == "(":
                depth += 1
            elif char == ")":
                depth -= 1
                if depth < 0:
                    return "Unbalanced closing parenthesis"
        index += 1

    if in_string:
        return "Unbalanced quote"
    if depth != 0:
        return "Unbalanced parentheses"
    return None
