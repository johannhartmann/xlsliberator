"""Tests for serious migration episodes and evidence-derived reporting."""

from __future__ import annotations

import json
import zipfile
from pathlib import Path

from click.testing import CliRunner

from xlsliberator.cli import cli
from xlsliberator.demo_corpus import (
    DemoCorpusManifest,
    DemoScenarioResult,
    generate_demo_corpus_report,
    search_demo_corpus,
)

ROOT = Path(__file__).parents[2]
MANIFEST_PATH = ROOT / "tests/corpus/manifests/episodes.json"


def test_serious_episode_layout_and_source_integrity() -> None:
    manifest = DemoCorpusManifest.load(MANIFEST_PATH)

    assert not manifest.verify(ROOT)
    assert len(manifest.episodes) == 8
    assert {episode.episode_id for episode in manifest.episodes} == {
        "interactive-game",
        "invoice-workflow",
        "database-dashboard",
        "addin-replacement",
        "legacy-xls-application",
        "xlsb-operations-model",
        "dependency-liberation",
        "hostile-workbook",
    }
    assert {episode.source.format for episode in manifest.episodes} == {
        "xls",
        "xlsx",
        "xlsm",
        "xlsb",
    }
    assert all(episode.target.status == "not_verified" for episode in manifest.episodes)
    assert all(episode.target.path is None for episode in manifest.episodes)


def test_search_index_and_subsets_are_generated_from_manifest() -> None:
    manifest = DemoCorpusManifest.load(MANIFEST_PATH)
    checked_index = json.loads(
        (ROOT / "tests/corpus/manifests/search-index.json").read_text(encoding="utf-8")
    )
    subsets = json.loads((ROOT / "tests/corpus/manifests/subsets.json").read_text(encoding="utf-8"))

    assert checked_index == manifest.search_index()
    for subset in ("pr", "nightly", "security"):
        assert subsets[subset] == sorted(
            episode.episode_id for episode in manifest.episodes if subset in episode.subsets
        )
    assert set(subsets["pr"]) < set(subsets["nightly"])
    assert {
        item["episode_id"]
        for item in search_demo_corpus(manifest, query="keyboard events", subset="nightly")
    } == {"interactive-game"}
    assert not search_demo_corpus(manifest, query="outlook", subset="security")


def test_public_scenarios_are_complete_and_hidden_tests_are_absent() -> None:
    manifest = DemoCorpusManifest.load(MANIFEST_PATH)
    public_index = json.loads(
        (ROOT / "tests/corpus/public-scenarios/index.json").read_text(encoding="utf-8")
    )

    assert public_index["hidden_acceptance_present"] is False
    assert set(public_index["episodes"]) == {episode.episode_id for episode in manifest.episodes}
    paths = [path.relative_to(ROOT).as_posix() for path in (ROOT / "tests/corpus").rglob("*")]
    assert not any("hidden" in Path(path).name.casefold() for path in paths)
    for episode in manifest.episodes:
        acceptance_text = (ROOT / episode.acceptance).read_text(encoding="utf-8").lower()
        assert "visibility: public" in acceptance_text
        assert "target_version: 26.2.4.2" in acceptance_text
        assert "ods file exists" not in acceptance_text


def test_hostile_workbook_is_inert_and_auditable() -> None:
    path = ROOT / "demos/hostile-workbook/source/HostileButInert.xlsx"

    with zipfile.ZipFile(path) as archive:
        names = set(archive.namelist())
        sheet = archive.read("xl/worksheets/sheet1.xml").decode()

    assert "xl/vbaProject.bin" not in names
    assert "Ignore all previous instructions" in sheet
    assert "169.254.169.254" in sheet
    assert "Do While True" in sheet
    assert "preserve as inert text" in sheet


def test_feature_report_preserves_not_measured_and_failures() -> None:
    manifest = DemoCorpusManifest.load(MANIFEST_PATH)
    not_run = generate_demo_corpus_report(manifest, [])
    passed = DemoScenarioResult(
        episode_id="legacy-xls-application",
        scenario_id="biff8-inventory",
        feature="biff8-import",
        source_format="xls",
        status="passed",
        evidence_path="artifacts/demo/legacy/biff8-inventory.json",
    )
    failed = DemoScenarioResult(
        episode_id="hostile-workbook",
        scenario_id="deny-network-process",
        feature="network-process-denial",
        source_format="xlsx",
        status="failed",
        evidence_path="artifacts/demo/hostile/network-process.json",
        failure_signature="security-network-boundary",
    )
    measured = generate_demo_corpus_report(manifest, [passed, failed])

    assert not_run.result_count == 0
    assert {item.capability_status for item in not_run.feature_status.values()} == {"not_measured"}
    assert measured.feature_status["biff8-import"].capability_status == "passed"
    assert measured.feature_status["network-process-denial"].capability_status == "failed"
    assert measured.format_status["xlsx"].capability_status == "failed"
    assert measured.format_status["xls"].capability_status == "passed"


def test_demo_corpus_cli_validates_searches_and_reports(tmp_path: Path) -> None:
    runner = CliRunner()
    report_path = tmp_path / "demo-report.json"

    validated = runner.invoke(
        cli,
        [
            "demo-corpus-validate",
            "--manifest",
            str(MANIFEST_PATH),
            "--search-index",
            str(ROOT / "tests/corpus/manifests/search-index.json"),
        ],
    )
    searched = runner.invoke(
        cli,
        [
            "demo-corpus-search",
            "--manifest",
            str(MANIFEST_PATH),
            "--query",
            "outlook",
            "--subset",
            "pr",
        ],
    )
    reported = runner.invoke(
        cli,
        [
            "demo-corpus-report",
            "--manifest",
            str(MANIFEST_PATH),
            "--results",
            str(ROOT / "tests/corpus/manifests/not-run-results.json"),
            "--output",
            str(report_path),
        ],
    )

    assert validated.exit_code == 0, validated.output
    assert "Validated 8 serious migration episodes" in validated.output
    assert searched.exit_code == 0, searched.output
    assert "invoice-workflow" in searched.output
    assert reported.exit_code == 0, reported.output
    assert json.loads(report_path.read_text(encoding="utf-8"))["result_count"] == 0
