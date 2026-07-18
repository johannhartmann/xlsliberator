"""Tests for evidence-derived capabilities and fail-closed release gates."""

from pathlib import Path

from click.testing import CliRunner

from xlsliberator.capability_matrix import (
    CapabilityMeasurement,
    ReleaseInputs,
    RuntimeEvidenceIdentity,
    generate_capability_report,
)
from xlsliberator.cli import cli
from xlsliberator.conformance_corpus import CorpusManifest
from xlsliberator.formula_corpus import (
    FormulaCorpusStatistics,
    collect_formula_corpus_statistics,
)

ROOT = Path(__file__).parents[2]


def _formula_corpus() -> FormulaCorpusStatistics:
    return collect_formula_corpus_statistics(
        ROOT / "tests" / "fixtures" / "formulas",
        display_path="tests/fixtures/formulas",
    )


def _runtime() -> RuntimeEvidenceIdentity:
    return RuntimeEvidenceIdentity(
        image_reference="xlsliberator-libreoffice:26.2.4.2",
        image_digest=f"sha256:{'a' * 64}",
        base_image_digest=f"sha256:{'e' * 64}",
        architecture="arm64",
        python_version="3.12.13",
        pyuno_identity={
            "uno_module_sha256": "b" * 64,
            "pyuno_native_sha256": "c" * 64,
        },
        office_binary_sha256="d" * 64,
        package_set=["libreoffice26.2=26.2.4.2-2", "libobasis26.2-pyuno=26.2.4.2-2"],
        runtime_variant="stock",
        source_commit="official-binary-distribution",
        patch_set_sha256="none",
    )


def _measurement(fixture_id: str, status: str = "passed") -> CapabilityMeasurement:
    return CapabilityMeasurement(
        evidence_id=f"evidence-{fixture_id}",
        fixture_id=fixture_id,
        source_format="xlsx",
        artifact_family="formula",
        scenario="recalculate",
        environment="docker-linux-arm64",
        runtime=_runtime() if status == "passed" else None,
        parse_coverage="passed",
        output_coverage="passed",
        target_runtime=status,
        source_differential="unavailable",
        evidence_bundle=f"evidence/{fixture_id}",
    )


def test_statuses_and_tiers_are_not_collapsed() -> None:
    manifest = CorpusManifest.load(ROOT / "corpus/manifest.json")
    measurements = [
        _measurement(fixture.fixture_id, "passed")
        if fixture.blocking
        else _measurement(fixture.fixture_id, "unavailable")
        for fixture in manifest.fixtures
    ]
    report = generate_capability_report(
        corpus=manifest,
        measurements=measurements,
        release_inputs=ReleaseInputs(
            p0_tests_passed=True,
            fail_open_paths_absent=True,
            source_artifacts_accounted=True,
            evidence_schemas_valid=True,
            security_suite_passed=True,
            agent_evaluation_passed=True,
        ),
        formula_corpus=_formula_corpus(),
    )

    assert report.release_ready
    assert report.summaries["target_runtime"].counts["unavailable"] > 0
    assert report.summaries["target_runtime"].counts["failed"] == 0
    assert report.summaries["source_differential"].decisive_pass_rate is None
    assert report.tier_counts["libreoffice-runtime-validated"] > 0


def test_missing_blocking_evidence_blocks_release() -> None:
    manifest = CorpusManifest.load(ROOT / "corpus/manifest.json")
    report = generate_capability_report(
        corpus=manifest,
        measurements=[],
        release_inputs=ReleaseInputs(
            p0_tests_passed=True,
            fail_open_paths_absent=True,
            source_artifacts_accounted=True,
            evidence_schemas_valid=True,
            security_suite_passed=True,
            agent_evaluation_passed=True,
        ),
        formula_corpus=_formula_corpus(),
    )

    assert not report.release_ready
    assert not next(gate for gate in report.release_gates if gate.name == "required-corpus").passed


def test_unsupported_blocking_fixture_is_not_green() -> None:
    manifest = CorpusManifest.load(ROOT / "corpus/manifest.json")
    measurements = [
        _measurement(fixture.fixture_id, "unsupported" if fixture.blocking else "unavailable")
        for fixture in manifest.fixtures
    ]
    report = generate_capability_report(
        corpus=manifest,
        measurements=measurements,
        release_inputs=ReleaseInputs(
            p0_tests_passed=True,
            fail_open_paths_absent=True,
            source_artifacts_accounted=True,
            evidence_schemas_valid=True,
            security_suite_passed=True,
            agent_evaluation_passed=True,
        ),
        formula_corpus=_formula_corpus(),
    )

    assert not report.release_ready
    gate = next(gate for gate in report.release_gates if gate.name == "required-corpus")
    assert not gate.passed
    assert "sample-formula-heavy-xlsx" in gate.reason


def test_required_corpus_reason_omits_green_blocking_fixtures() -> None:
    manifest = CorpusManifest.load(ROOT / "corpus/manifest.json")
    green_fixture = next(fixture.fixture_id for fixture in manifest.fixtures if fixture.blocking)
    report = generate_capability_report(
        corpus=manifest,
        measurements=[_measurement(green_fixture)],
        release_inputs=ReleaseInputs(
            p0_tests_passed=True,
            fail_open_paths_absent=True,
            source_artifacts_accounted=True,
            evidence_schemas_valid=True,
            security_suite_passed=True,
            agent_evaluation_passed=True,
        ),
        formula_corpus=_formula_corpus(),
    )

    gate = next(gate for gate in report.release_gates if gate.name == "required-corpus")
    assert not gate.passed
    assert green_fixture not in gate.reason


def test_cli_writes_generated_reports_and_blocks_release(tmp_path: Path) -> None:
    evidence = tmp_path / "evidence.json"
    evidence.write_text('{"schema_version":"1.0.0","measurements":[]}', encoding="utf-8")
    release_inputs = tmp_path / "release.json"
    release_inputs.write_text(
        """{
          "p0_tests_passed": true,
          "fail_open_paths_absent": true,
          "source_artifacts_accounted": true,
          "evidence_schemas_valid": true,
          "security_suite_passed": true
        }""",
        encoding="utf-8",
    )
    json_output = tmp_path / "matrix.json"
    markdown_output = tmp_path / "matrix.md"

    result = CliRunner().invoke(
        cli,
        [
            "capability-report",
            "--corpus",
            str(ROOT / "corpus/manifest.json"),
            "--evidence",
            str(evidence),
            "--release-inputs",
            str(release_inputs),
            "--json-output",
            str(json_output),
            "--markdown-output",
            str(markdown_output),
            "--check-release",
        ],
    )

    assert result.exit_code == 1
    assert '"generated_from_evidence": true' in json_output.read_text(encoding="utf-8")
    assert "Do not edit by hand" in markdown_output.read_text(encoding="utf-8")
