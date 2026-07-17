"""Evidence-conservative statistics for the deterministic formula corpus."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Literal

from pydantic import BaseModel, ConfigDict, Field

from xlsliberator.formula_rules import FormulaRuleRegistry


class FormulaCorpusStatistics(BaseModel):
    """Versioned statistics that never infer runtime support from fixtures."""

    model_config = ConfigDict(extra="forbid")

    schema_version: Literal["1.0.0"] = "1.0.0"
    corpus_path: str
    minimized_regression_fixtures: int = 0
    registered_rules: int = 0
    covered_rules: int = 0
    uncovered_rules: list[str] = Field(default_factory=list)
    fixture_rule_counts: dict[str, int] = Field(default_factory=dict)
    source_differential_status: Literal["not_measured"] = "not_measured"
    source_differential_passed: int = 0
    source_differential_failed: int = 0


def collect_formula_corpus_statistics(
    corpus_dir: Path,
    registry: FormulaRuleRegistry | None = None,
    *,
    display_path: str | None = None,
) -> FormulaCorpusStatistics:
    """Count minimized rule fixtures without promoting them to runtime evidence."""
    active_registry = registry or FormulaRuleRegistry.with_default_rules()
    manifest_rules = active_registry.manifest().get("rules")
    if not isinstance(manifest_rules, list):
        raise ValueError("formula rule registry manifest has no rule list")
    registered = {
        str(item["name"]) for item in manifest_rules if isinstance(item, dict) and "name" in item
    }
    counts: dict[str, int] = {}
    fixture_count = 0
    for path in sorted(corpus_dir.glob("*.json")):
        payload: dict[str, Any] = json.loads(path.read_text(encoding="utf-8"))
        rule = payload.get("rule")
        if not isinstance(rule, str):
            continue
        fixture_count += 1
        counts[rule] = counts.get(rule, 0) + 1
    covered = registered & set(counts)
    return FormulaCorpusStatistics(
        corpus_path=display_path or str(corpus_dir),
        minimized_regression_fixtures=fixture_count,
        registered_rules=len(registered),
        covered_rules=len(covered),
        uncovered_rules=sorted(registered - covered),
        fixture_rule_counts=dict(sorted(counts.items())),
    )
