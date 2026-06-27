"""Deterministic formula repair loop skeleton."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from xlsliberator.formula_rules import FormulaRuleRegistry
from xlsliberator.validation_models import SourceRef, TargetRef, WorkbookArtifactIR


@dataclass(frozen=True)
class FormulaRepairEvidence:
    """Evidence that a formula may need repair."""

    source_ref: SourceRef
    target_ref: TargetRef
    source_formula: str
    target_formula: str
    parse_error: str | None = None
    value_mismatch: str | None = None


@dataclass(frozen=True)
class FormulaRepairAttempt:
    """One deterministic repair attempt."""

    evidence: FormulaRepairEvidence
    rule_name: str
    before: str
    after: str
    success: bool
    error: str | None = None


class FormulaRepairLoop:
    """Repair loop driven by deterministic formula rules."""

    def __init__(self, registry: FormulaRuleRegistry | None = None) -> None:
        self.registry = registry or FormulaRuleRegistry.with_default_rules()

    def collect_evidence(
        self, evidence: list[FormulaRepairEvidence]
    ) -> list[FormulaRepairEvidence]:
        """Return provided evidence until target-backed collection is implemented."""
        return evidence

    def collect_evidence_from_inventory(
        self, inventory: WorkbookArtifactIR
    ) -> list[FormulaRepairEvidence]:
        """Build repair evidence from source formulas that match a repair rule.

        Only formulas a deterministic rule can act on become evidence, so the
        repair gate exercises real formulas instead of an always-empty list.
        """
        evidence: list[FormulaRepairEvidence] = []
        for formula in inventory.formulas:
            if not self.registry.matching_rules(formula.formula_text):
                continue
            source_ref = formula.source_ref
            evidence.append(
                FormulaRepairEvidence(
                    source_ref=source_ref,
                    target_ref=TargetRef(
                        target_file=source_ref.source_file,
                        sheet=source_ref.sheet,
                        cell_range=source_ref.cell_range,
                        artifact_type="formula",
                        artifact_id=source_ref.artifact_id,
                    ),
                    source_formula=formula.formula_text,
                    target_formula=formula.formula_text,
                )
            )
        return evidence

    def propose_repairs(self, evidence: list[FormulaRepairEvidence]) -> list[FormulaRepairAttempt]:
        """Create repair attempts for evidence with matching deterministic rules."""
        attempts: list[FormulaRepairAttempt] = []
        for item in evidence:
            result = self.registry.apply_first(item.target_formula)
            if result is None:
                attempts.append(
                    FormulaRepairAttempt(
                        evidence=item,
                        rule_name="unresolved",
                        before=item.target_formula,
                        after=item.target_formula,
                        success=False,
                        error="No deterministic rule matched",
                    )
                )
                continue
            attempts.append(
                FormulaRepairAttempt(
                    evidence=item,
                    rule_name=result.rule_name,
                    before=result.before,
                    after=result.after,
                    success=result.success,
                    error=result.error,
                )
            )
        return attempts

    def apply_repairs_to_ods(
        self,
        ods_path: Path,
        attempts: list[FormulaRepairAttempt],
    ) -> list[FormulaRepairAttempt]:
        """Placeholder for future ODS patching; returns reportable attempts."""
        _ = ods_path
        return attempts
