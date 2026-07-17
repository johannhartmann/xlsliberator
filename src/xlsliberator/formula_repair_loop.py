"""Evidence-driven deterministic formula repair planning."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from xlsliberator.formula_certification import FormulaCertificationResult
from xlsliberator.formula_rules import FormulaRuleRegistry
from xlsliberator.validation_models import SourceRef, TargetRef


@dataclass(frozen=True)
class FormulaRepairEvidence:
    """Evidence that a formula may need repair."""

    source_ref: SourceRef
    target_ref: TargetRef
    source_formula: str
    target_formula: str
    target_parser_result: dict[str, Any] | None = None
    source_value: dict[str, Any] | None = None
    target_value: dict[str, Any] | None = None
    dependencies: tuple[str, ...] = ()
    spill_context: dict[str, Any] | None = None
    scenario_id: str | None = None
    observation_id: str | None = None
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

    def collect_evidence_from_result(
        self,
        result: FormulaCertificationResult,
        *,
        scenario_id: str | None = None,
    ) -> list[FormulaRepairEvidence]:
        """Convert actual failed parser/trace records into deterministic repair input."""
        evidence: list[FormulaRepairEvidence] = []
        for record in result.records:
            if record.status.value == "passed" or record.target_ref is None:
                continue
            if not record.target_formula:
                continue
            evidence.append(
                FormulaRepairEvidence(
                    source_ref=record.source_ref,
                    target_ref=record.target_ref,
                    source_formula=record.source_formula,
                    target_formula=record.target_formula,
                    target_parser_result=record.target_parser_result,
                    source_value=(
                        record.source_observation.model_dump(mode="json")
                        if record.source_observation
                        else None
                    ),
                    target_value=(
                        record.target_observation.model_dump(mode="json")
                        if record.target_observation
                        else None
                    ),
                    dependencies=tuple(record.dependencies),
                    spill_context=dict(record.spill_context),
                    scenario_id=scenario_id,
                    observation_id=record.observation_id,
                    parse_error="; ".join(
                        error for error in record.errors if "Parser" in error or "parse" in error
                    )
                    or None,
                    value_mismatch="; ".join(
                        error for error in record.errors if "runtime value/error" in error
                    )
                    or None,
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

    @staticmethod
    def worker_formula_repairs(
        attempts: list[FormulaRepairAttempt],
    ) -> list[dict[str, str]]:
        """Return only deterministic, located repairs suitable for one transaction."""
        repairs: list[dict[str, str]] = []
        for attempt in attempts:
            target_ref = attempt.evidence.target_ref
            if not attempt.success or not target_ref.sheet or not target_ref.cell_range:
                continue
            repairs.append(
                {
                    "sheet": target_ref.sheet,
                    "address": target_ref.cell_range,
                    "formula": attempt.after,
                    "rule_name": attempt.rule_name,
                }
            )
        return repairs
