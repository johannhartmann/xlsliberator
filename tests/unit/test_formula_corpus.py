"""Formula corpus statistics and capability-matrix synchronization tests."""

import json
from pathlib import Path

from xlsliberator.formula_corpus import collect_formula_corpus_statistics


def test_formula_corpus_statistics_match_capability_matrix() -> None:
    root = Path(__file__).parents[2]
    statistics = collect_formula_corpus_statistics(
        root / "tests" / "fixtures" / "formulas",
        display_path="tests/fixtures/formulas",
    )
    capability_matrix = json.loads(
        (root / "docs" / "capability_matrix.json").read_text(encoding="utf-8")
    )

    assert statistics.model_dump(mode="json") == capability_matrix["formula_corpus"]
    assert statistics.covered_rules == statistics.registered_rules
    assert statistics.source_differential_status == "not_measured"
