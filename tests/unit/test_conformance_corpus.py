"""Tests for corpus truthfulness, fuzzing, and failure minimization."""

from __future__ import annotations

import json
from pathlib import Path

from xlsliberator.conformance_corpus import (
    CorpusExecution,
    CorpusManifest,
    check_metamorphic_relations,
    copy_move_rename_recipe,
    corpus_statistics,
    corpus_trend_report,
    differential_fuzz,
    generate_charts_pivots_controls_recipe,
    generate_environment_recipe,
    generate_formula_cases,
    generate_names_tables_recipe,
    generate_styles_validations_recipe,
    generate_vba_recipe,
    normalized_failure_signature,
    recipe_save_reopen_stable,
)
from xlsliberator.failure_minimizer import WorkbookCandidate, minimize_failure

ROOT = Path(__file__).parents[2]


def test_corpus_manifest_is_non_confidential_and_integrity_checked() -> None:
    manifest = CorpusManifest.load(ROOT / "corpus/manifest.json")

    assert not manifest.verify_files(ROOT)
    assert {fixture.format for fixture in manifest.fixtures} >= {"xls", "xlsx", "xlsm", "xlsb"}
    assert {fixture.origin for fixture in manifest.fixtures} >= {
        "generated",
        "public",
        "malicious",
        "regression",
    }
    assert all(not fixture.confidential for fixture in manifest.fixtures)


def test_release_certification_fixtures_describe_real_target_artifacts() -> None:
    manifest = CorpusManifest.load(ROOT / "corpus/manifest.json")
    by_id = {fixture.fixture_id: fixture for fixture in manifest.fixtures}

    vba_fixture = by_id["sample-vba-workbook"]
    generated_vba_fixture = by_id["generated-vba-xlsm"]
    controls_fixture = by_id["sample-controls-events-workbook"]
    assert generated_vba_fixture.expected[0].assertions == {
        "actual_xlsm_vba_project": True,
        "target_compatibility_execution": True,
    }
    assert vba_fixture.format == "xlsm"
    assert vba_fixture.expected[0].expected_status == "passed"
    assert vba_fixture.expected[0].assertions["actual_xlsm_vba_project"] is True
    assert controls_fixture.format == "ods"
    assert controls_fixture.materialization == "generated"
    assert controls_fixture.expected[0].expected_status == "passed"


def test_differential_fuzz_keeps_unavailable_distinct() -> None:
    cases = generate_formula_cases(seed=16, count=4)

    unavailable = differential_fuzz(cases, source=None, target=lambda formula: formula)
    passed = differential_fuzz(
        cases, source=lambda formula: formula, target=lambda formula: formula
    )

    assert {result.status for result in unavailable} == {"unavailable"}
    assert {result.status for result in passed} == {"passed"}


def test_all_required_generator_families_are_deterministic() -> None:
    recipes = [
        generate_names_tables_recipe(1),
        generate_styles_validations_recipe(2),
        generate_charts_pivots_controls_recipe(3),
        generate_vba_recipe(4),
        generate_environment_recipe(
            5, locale="de-DE", date_system="1904", calculation_mode="manual"
        ),
    ]

    assert len({recipe.canonical_sha256 for recipe in recipes}) == len(recipes)
    assert all(recipe_save_reopen_stable(recipe) for recipe in recipes)
    assert recipes[2].format == "xlsb"
    assert recipes[3].format == "xlsm"


def test_metamorphic_rewrites_rename_and_independent_execution() -> None:
    case = generate_formula_cases(seed=9, count=1)[0]

    results = check_metamorphic_relations(
        case,
        execute=lambda _formula: "42",
        independent_target=lambda _formula: "42",
    )
    renamed = copy_move_rename_recipe(
        generate_names_tables_recipe(9), source_sheet="Input", target_sheet="Source"
    )

    assert {result.status for result in results} == {"passed"}
    assert "Source" in renamed.sheets
    assert all("Input." not in reference for reference in renamed.names.values())


def test_failure_signatures_are_normalized_and_deduplicated() -> None:
    first = normalized_failure_signature(
        gate="formula", error_type="mismatch", trace_diff=["/tmp123/job/cell42"]
    )
    second = normalized_failure_signature(
        gate="formula", error_type="mismatch", trace_diff=["/tmp999/job/cell77"]
    )
    manifest = CorpusManifest.load(ROOT / "corpus/manifest.json")
    executions = [
        CorpusExecution(
            fixture_id="a",
            scenario="formula",
            environment="docker",
            target="libreoffice",
            target_version="26.2.4.2",
            status="failed",
            failure_signature=first,
            evidence_path="evidence/a.json",
        ),
        CorpusExecution(
            fixture_id="b",
            scenario="formula",
            environment="docker",
            target="libreoffice",
            target_version="26.2.4.2",
            status="failed",
            failure_signature=second,
            evidence_path="evidence/b.json",
        ),
        CorpusExecution(
            fixture_id="c",
            scenario="formula",
            environment="docker",
            target="libreoffice",
            target_version="26.2.4.2",
            status="skipped",
        ),
    ]

    statistics = corpus_statistics(manifest, executions)

    assert first == second
    assert statistics.status_counts["failed"] == 2
    assert statistics.status_counts["skipped"] == 1
    assert statistics.unique_failure_signatures == 1
    assert statistics.duplicate_failures == 1

    trend = corpus_trend_report(manifest, executions, previous=statistics)
    assert trend.fixture_delta == 0
    assert set(trend.status_deltas) == {
        "passed",
        "failed",
        "skipped",
        "unavailable",
        "unsupported",
        "waived",
    }
    assert all(delta == 0 for delta in trend.status_deltas.values())


def test_multifeature_failure_is_automatically_minimized() -> None:
    fixture = json.loads(
        (ROOT / "tests/fixtures/corpus/multifeature_failure.json").read_text(encoding="utf-8")
    )
    original = WorkbookCandidate.model_validate(fixture["original"])
    required = WorkbookCandidate.model_validate(fixture["minimal_requirements"])
    signature = str(fixture["failure_signature"])

    def predicate(candidate: WorkbookCandidate) -> str | None:
        for field in (
            "sheets",
            "ranges",
            "formulas",
            "vba_modules",
            "vba_procedures",
            "package_parts",
        ):
            if not set(getattr(required, field)).issubset(getattr(candidate, field)):
                return None
        return signature

    evidence = minimize_failure(
        original,
        expected_signature=signature,
        predicate=predicate,
    )

    assert evidence.minimized == required
    assert evidence.minimized.size < evidence.original.size
    assert any(step.retained_failure for step in evidence.steps)
    regression = json.loads(
        (ROOT / "tests/fixtures/corpus/tdf_172479_minimized.json").read_text(encoding="utf-8")
    )
    preserved = json.loads((ROOT / regression["minimization_evidence"]).read_text(encoding="utf-8"))
    assert WorkbookCandidate.model_validate(regression["workbook"]) == evidence.minimized
    assert preserved["failure_signature"] == evidence.failure_signature
    assert preserved["original_size"] == evidence.original.size
    assert preserved["minimized_size"] == evidence.minimized.size
