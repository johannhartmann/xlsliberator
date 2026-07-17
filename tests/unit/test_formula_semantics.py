"""Semantic formula IR, dialect, feature, and generated-property tests."""

import json
from pathlib import Path

import pytest
from hypothesis import given
from hypothesis import strategies as st

from xlsliberator.formula_engine import FormulaDialect
from xlsliberator.formula_rules import FormulaRuleRegistry
from xlsliberator.formula_semantics import build_formula_ir
from xlsliberator.validation_models import SourceRef


def _ref() -> SourceRef:
    return SourceRef(
        source_file="book.xlsx",
        sheet="Data",
        cell_range="C3",
        artifact_type="formula",
        artifact_id="formula:Data!C3",
    )


@pytest.mark.parametrize(
    ("dialect", "formula", "dependency"),
    [
        (FormulaDialect.EXCEL_A1, "=SUM(A1:B2)", "A1:B2"),
        (FormulaDialect.EXCEL_R1C1, "=SUM(R[-1]C:R[2]C[1])", "R[-1]C:R[2]C[1]"),
        (FormulaDialect.CALC_A1, "=SUM(A1:B2)", "A1:B2"),
        (FormulaDialect.OPENFORMULA, "of:=SUM([.A1:.B2])", "A1"),
    ],
)
def test_formula_ir_supports_all_declared_dialects(
    dialect: FormulaDialect,
    formula: str,
    dependency: str,
) -> None:
    artifact = build_formula_ir(source_ref=_ref(), formula=formula, dialect=dialect)

    assert artifact.schema_version == "2.0.0"
    assert artifact.original_formula_text == formula
    assert artifact.sheet_context == "Data"
    assert artifact.cell_context == "C3"
    assert artifact.dialect == dialect.value
    assert any(dependency in item for item in artifact.dependencies)


def test_formula_ir_accounts_for_advanced_semantic_features() -> None:
    artifact = build_formula_ir(
        source_ref=_ref(),
        formula="=LET(x,FILTER(Table1[Amount],Table1[Amount]>0),LAMBDA(y,y+Sheet1:Sheet3!A1)(@x#))",
        dialect=FormulaDialect.EXCEL_A1,
        array_metadata={"spill_range": "C3:C9", "dynamic": True},
        calculation_settings={"mode": "manual", "iterate": True},
        calculation_order={"declared_chain_position": 7},
    )

    assert {
        "structured_reference",
        "dynamic_array",
        "implicit_intersection",
        "spill_reference",
        "let",
        "lambda",
        "3d_reference",
    } <= set(artifact.semantic_features)
    assert artifact.array_metadata["spill_range"] == "C3:C9"
    assert artifact.calculation_settings == {"mode": "manual", "iterate": True}
    assert artifact.calculation_order == {"declared_chain_position": 7}
    assert artifact.unsupported_reasons == []
    assert artifact.semantic_diagnostics
    assert artifact.runtime_evidence_requirements


def test_formula_ir_detects_volatile_and_external_formula_context() -> None:
    artifact = build_formula_ir(
        source_ref=_ref(),
        formula="=NOW()+OFFSET([Budget.xlsx]Data!A1,1,0)",
    )

    assert artifact.volatility_flags == ["NOW", "OFFSET"]
    assert "external_reference" in artifact.semantic_features


def test_formula_ir_uses_extracted_array_formula_metadata() -> None:
    artifact = build_formula_ir(
        source_ref=_ref(),
        formula="=SUM(A1:A3*B1:B3)",
        array_metadata={"formula_type": "ArrayFormula", "array_range": "C3:C5"},
    )

    assert "array_formula" in artifact.semantic_features
    assert artifact.array_metadata["array_range"] == "C3:C5"


@given(
    left=st.integers(min_value=1, max_value=9999),
    right=st.integers(min_value=1, max_value=9999),
    operator=st.sampled_from(["+", "-", "*", "/"]),
)
def test_generated_arithmetic_formula_ir_is_deterministic(
    left: int,
    right: int,
    operator: str,
) -> None:
    formula = f"={left}{operator}{right}"

    first = build_formula_ir(source_ref=_ref(), formula=formula)
    second = build_formula_ir(source_ref=_ref(), formula=formula)

    assert first.model_dump(mode="json") == second.model_dump(mode="json")
    assert first.tokens is not None
    assert first.ast == {"kind": "token_stream", "token_count": len(first.tokens)}


def test_minimized_fixture_covers_each_registered_rule() -> None:
    fixture_dir = Path(__file__).parents[1] / "fixtures" / "formulas"
    fixtures = {
        payload["rule"]: payload
        for path in fixture_dir.glob("*.json")
        if (payload := json.loads(path.read_text(encoding="utf-8")))
    }
    registry = FormulaRuleRegistry.with_default_rules()

    assert set(fixtures) == {item["name"] for item in registry.manifest()["rules"]}  # type: ignore[index]
    for rule_name, fixture in fixtures.items():
        result = registry.apply_first(str(fixture["source"]))
        assert result is not None
        assert result.rule_name == rule_name
        assert result.success
        assert result.after == fixture["expected"]
