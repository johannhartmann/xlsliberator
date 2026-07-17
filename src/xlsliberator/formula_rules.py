"""Formula repair rule registry."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Protocol

from xlsliberator.formula_ast_transformer import FormulaASTTransformer, FormulaTransformError

FORMULA_RULE_REGISTRY_VERSION = "1.0.0"


@dataclass(frozen=True)
class RuleMatch:
    """Formula rule match."""

    rule_name: str
    formula: str
    details: dict[str, str]


@dataclass(frozen=True)
class RuleApplicationResult:
    """Formula rule application result."""

    rule_name: str
    before: str
    after: str
    success: bool
    error: str | None = None


class FormulaRule(Protocol):
    """Formula repair rule protocol."""

    name: str

    def match(self, formula: str) -> RuleMatch | None:
        """Return a match when the rule applies."""

    def apply(self, formula: str) -> RuleApplicationResult:
        """Apply the rule to a formula."""


class IndirectAddressRule:
    """Existing deterministic INDIRECT/ADDRESS repair as a registry rule."""

    name = "indirect_address"
    version = "1.0.0"

    def __init__(self, sheet_mapping: dict[str, str] | None = None) -> None:
        self.sheet_mapping = sheet_mapping or {}

    def match(self, formula: str) -> RuleMatch | None:
        """Match formulas containing INDIRECT and ADDRESS."""
        if "INDIRECT" in formula.upper() and "ADDRESS" in formula.upper():
            return RuleMatch(self.name, formula, {})
        return None

    def apply(self, formula: str) -> RuleApplicationResult:
        """Apply the existing AST transformer."""
        try:
            after = FormulaASTTransformer(self.sheet_mapping).transform_indirect_address_to_offset(
                formula
            )
            return RuleApplicationResult(self.name, formula, after, True)
        except FormulaTransformError as exc:
            return RuleApplicationResult(self.name, formula, formula, False, str(exc))


class FormulaRuleRegistry:
    """Registry for deterministic formula repair rules."""

    def __init__(self, rules: list[FormulaRule] | None = None) -> None:
        self.schema_version = FORMULA_RULE_REGISTRY_VERSION
        self.rules = rules or []

    @classmethod
    def with_default_rules(
        cls,
        sheet_mapping: dict[str, str] | None = None,
    ) -> FormulaRuleRegistry:
        """Return registry with built-in rules."""
        return cls([IndirectAddressRule(sheet_mapping)])

    def matching_rules(self, formula: str) -> list[FormulaRule]:
        """Return rules matching a formula."""
        return [rule for rule in self.rules if rule.match(formula)]

    def apply_first(self, formula: str) -> RuleApplicationResult | None:
        """Apply the first matching rule, if any."""
        for rule in self.rules:
            if rule.match(formula):
                return rule.apply(formula)
        return None

    def manifest(self) -> dict[str, object]:
        """Export stable rule identities for evidence and capability reporting."""
        return {
            "schema_version": self.schema_version,
            "rules": [
                {"name": rule.name, "version": str(getattr(rule, "version", "unversioned"))}
                for rule in self.rules
            ],
        }
