"""Dialect-aware semantic formula analysis without claiming target execution."""

from __future__ import annotations

import re
from typing import Any

from openpyxl.formula import Tokenizer

from xlsliberator.formula_engine import FormulaDialect
from xlsliberator.validation_models import FormulaIR, SourceRef, TargetRef

FORMULA_SEMANTICS_VERSION = "1.0.0"

_VOLATILE_FUNCTIONS = {
    "CELL",
    "INDIRECT",
    "INFO",
    "NOW",
    "OFFSET",
    "RAND",
    "RANDBETWEEN",
    "TODAY",
}
_DYNAMIC_ARRAY_FUNCTIONS = {
    "CHOOSECOLS",
    "CHOOSEROWS",
    "DROP",
    "EXPAND",
    "FILTER",
    "HSTACK",
    "MAKEARRAY",
    "RANDARRAY",
    "SEQUENCE",
    "SORT",
    "SORTBY",
    "TAKE",
    "TOCOL",
    "TOROW",
    "UNIQUE",
    "VSTACK",
    "WRAPCOLS",
    "WRAPROWS",
}
_FUNCTION_RE = re.compile(r"(?i)(?<![A-Z0-9_.])([A-Z_][A-Z0-9_.]*)\s*\(")
_A1_REFERENCE_RE = re.compile(
    r"(?i)(?<![A-Z0-9_])(?:'[^']+'|[A-Z_][A-Z0-9_. ]*)?!?"
    r"\$?[A-Z]{1,3}\$?[1-9][0-9]*(?::\$?[A-Z]{1,3}\$?[1-9][0-9]*)?"
)
_R1C1_REFERENCE_RE = re.compile(
    r"(?i)(?<![A-Z0-9_])R(?:\[-?\d+\]|-?\d+)?C(?:\[-?\d+\]|-?\d+)?"
    r"(?::R(?:\[-?\d+\]|-?\d+)?C(?:\[-?\d+\]|-?\d+)?)?"
)


def build_formula_ir(
    *,
    source_ref: SourceRef,
    formula: str,
    dialect: FormulaDialect = FormulaDialect.EXCEL_A1,
    target_ref: TargetRef | None = None,
    target_formula: str | None = None,
    name_context: str | None = None,
    array_metadata: dict[str, Any] | None = None,
    calculation_settings: dict[str, Any] | None = None,
    calculation_order: dict[str, Any] | None = None,
) -> FormulaIR:
    """Build conservative semantic metadata for one formula representation."""
    tokens, ast, token_error = _tokenize(formula, dialect)
    functions = sorted({match.upper() for match in _FUNCTION_RE.findall(formula)})
    features = _semantic_features(formula, functions)
    array_context = array_metadata or {}
    if array_context.get("array_range") or str(array_context.get("formula_type")) == "ArrayFormula":
        features.append("array_formula")
    if array_context.get("spill_range") or array_context.get("dynamic"):
        features.append("dynamic_array")
    features = sorted(set(features))
    if name_context:
        features.append("defined_name_formula")
        if re.search(rf"(?i)(?<![A-Z0-9_]){re.escape(name_context)}(?![A-Z0-9_])", formula):
            features.append("recursive_name")
        features = sorted(set(features))
    diagnostics = [f"semantic tokenization failed: {token_error}"] if token_error else []
    return FormulaIR(
        source_ref=source_ref,
        target_ref=target_ref,
        formula_text=formula,
        original_formula_text=formula,
        target_formula_text=target_formula,
        dialect=dialect.value,
        tokens=tokens,
        ast=ast,
        sheet_context=source_ref.sheet,
        cell_context=source_ref.cell_range,
        name_context=name_context,
        dependencies=_dependencies(formula, dialect),
        volatility_flags=sorted(set(functions) & _VOLATILE_FUNCTIONS),
        semantic_features=features,
        semantic_diagnostics=diagnostics,
        runtime_evidence_requirements=_runtime_evidence_requirements(features),
        array_metadata=array_context,
        calculation_settings=calculation_settings or {},
        calculation_order=calculation_order or {},
        unsupported_reasons=[],
    )


def _tokenize(
    formula: str, dialect: FormulaDialect
) -> tuple[list[dict[str, str]] | None, dict[str, Any] | None, str | None]:
    if dialect not in {FormulaDialect.EXCEL_A1, FormulaDialect.CALC_A1}:
        return None, {"kind": "formula", "dialect": dialect.value}, None
    candidate = formula
    if dialect is FormulaDialect.CALC_A1:
        candidate = candidate.replace(";", ",")
    try:
        items = Tokenizer(candidate).items
    except Exception as exc:
        return None, None, f"{type(exc).__name__}: {exc}"
    tokens = [
        {"value": str(item.value), "type": str(item.type), "subtype": str(item.subtype)}
        for item in items
    ]
    return tokens, {"kind": "token_stream", "token_count": len(tokens)}, None


def _dependencies(formula: str, dialect: FormulaDialect) -> list[str]:
    pattern = _R1C1_REFERENCE_RE if dialect is FormulaDialect.EXCEL_R1C1 else _A1_REFERENCE_RE
    return sorted({match.group(0) for match in pattern.finditer(formula)})


def _semantic_features(formula: str, functions: list[str]) -> list[str]:
    upper = formula.upper()
    features: set[str] = set()
    if "[" in formula and "]" in formula:
        features.add("structured_reference")
    if re.search(r"\[[^\]]+\.(?:XLSX?|XLSM|XLSB|ODS)\]", formula, re.IGNORECASE):
        features.add("external_reference")
    if re.search(r"(?:'[^']+'|[A-Za-z0-9_]+):(?:'[^']+'|[A-Za-z0-9_]+)!", formula):
        features.add("3d_reference")
    if "@" in formula:
        features.add("implicit_intersection")
    if "#" in formula:
        features.add("spill_reference")
    if set(functions) & _DYNAMIC_ARRAY_FUNCTIONS:
        features.add("dynamic_array")
    if formula.startswith("{") and formula.endswith("}"):
        features.add("array_formula")
    if "LET" in functions:
        features.add("let")
    if "LAMBDA" in functions:
        features.add("lambda")
    if "_XLPM." in upper or "_XLFN." in upper:
        features.add("future_function")
    return sorted(features)


def _runtime_evidence_requirements(features: list[str]) -> list[str]:
    return [
        f"{feature} requires source/target runtime evidence"
        for feature in (
            "3d_reference",
            "dynamic_array",
            "external_reference",
            "lambda",
            "recursive_name",
            "structured_reference",
        )
        if feature in features
    ]
