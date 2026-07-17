"""Tests for formula rule registry."""

from xlsliberator.formula_rules import FormulaRuleRegistry, IndirectAddressRule


def test_indirect_address_rule_detection() -> None:
    """INDIRECT/ADDRESS formulas should match the default rule."""
    rule = IndirectAddressRule()

    assert rule.match('=INDIRECT(ADDRESS(1;2;4;1;"Sheet2"))') is not None
    assert rule.match("=SUM(A1:A2)") is None


def test_formula_rule_application_noop_when_no_match() -> None:
    """Registry should return None for formulas without matching rules."""
    registry = FormulaRuleRegistry.with_default_rules()

    assert registry.apply_first("=SUM(A1:A2)") is None


def test_indirect_address_rule_preserves_existing_expected_output() -> None:
    """Default registry should keep existing transformer behavior."""
    registry = FormulaRuleRegistry.with_default_rules()
    result = registry.apply_first('=INDIRECT(ADDRESS(1;2;4;1;"Sheet2"))')

    assert result is not None
    assert result.success
    assert result.after == '=INDIRECT("Sheet2!"&ADDRESS(1;2;4;1))'
