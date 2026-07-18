"""Public acceptance, evidence, and mutation-system tests."""

from __future__ import annotations

import hashlib
import json
import zipfile
from pathlib import Path

import pytest
from click.testing import CliRunner
from pydantic import ValidationError

from xlsliberator.migration_check import (
    build_migration_report,
    cli,
    load_acceptance,
    run_acceptance,
)
from xlsliberator.scenarios.acceptance_evidence import inspect_acceptance_evidence
from xlsliberator.scenarios.assertions import evaluate_trace
from xlsliberator.scenarios.models import (
    AcceptanceDefinition,
    Action,
    ActionKind,
    ComparisonRules,
    EnvironmentManifest,
    MigrationMetadata,
    ObservationKind,
    ObservationRequest,
    ObservationValue,
    RuntimeTrace,
    Scenario,
    ScenarioStep,
    ValueKind,
)
from xlsliberator.scenarios.mutation import run_mutation_campaign
from xlsliberator.scenarios.runner import FakeScenarioRunner
from xlsliberator.validation_models import GateExecutionStatus

MIMETYPE = "application/vnd.oasis.opendocument.spreadsheet"
MANIFEST_NS = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"


class FakePathRunner:
    """Deterministic path-based adapter used in place of the Prompt 06 service."""

    def __init__(
        self,
        observations: dict[str, object | ObservationValue],
        *,
        statuses: dict[str, GateExecutionStatus] | None = None,
    ) -> None:
        self.observations = observations
        self.statuses = statuses

    def run(
        self,
        workbook: Path,
        environment: EnvironmentManifest,
        scenario: Scenario,
    ) -> RuntimeTrace:
        return FakeScenarioRunner(
            "fake_target",
            self.observations,
            statuses=self.statuses,
        ).run(workbook.read_bytes(), environment, scenario)


def _acceptance(*, tolerance: float = 0.0) -> AcceptanceDefinition:
    return AcceptanceDefinition(
        migration=MigrationMetadata(
            id="invoice-migration",
            title="Invoice migration",
            target_workbook="target.ods",
            authored_by="author@example.test",
            reviewed_by="reviewer@example.test",
            requirements=["The total remains 10."],
        ),
        scenario=Scenario(
            id="public-acceptance",
            steps=[
                ScenarioStep(
                    id="open",
                    action=Action(kind=ActionKind.OPEN),
                    observations_after=[
                        ObservationRequest(
                            id="total",
                            kind=ObservationKind.CELL_VALUE,
                            selector={"sheet": "Sheet1", "address": "A1"},
                            expected=ObservationValue(kind=ValueKind.NUMBER, value=10.0),
                            comparison=ComparisonRules(absolute_tolerance=tolerance),
                        )
                    ],
                )
            ],
        ),
    )


def _write_ods(path: Path) -> Path:
    content = """<?xml version="1.0" encoding="UTF-8"?>
<office:document-content
 xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
 xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
 <office:body><office:spreadsheet><table:table table:name="Sheet1">
  <table:table-row><table:table-cell table:formula="of:=1+1"/></table:table-row>
 </table:table></office:spreadsheet></office:body>
</office:document-content>"""
    manifest = f"""<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="{MANIFEST_NS}" manifest:version="1.3">
 <manifest:file-entry manifest:full-path="/" manifest:media-type="{MIMETYPE}"/>
 <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
 <manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>
 <manifest:file-entry manifest:full-path="Scripts/python/" manifest:media-type="application/binary"/>
 <manifest:file-entry manifest:full-path="Scripts/python/main.py" manifest:media-type="application/binary"/>
</manifest:manifest>"""
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("mimetype", MIMETYPE, compress_type=zipfile.ZIP_STORED)
        archive.writestr("content.xml", content)
        archive.writestr("styles.xml", "<styles/>")
        archive.writestr("Scripts/python/main.py", "def total():\n    return 1\n")
        archive.writestr("META-INF/manifest.xml", manifest)
    return path


def test_yaml_models_cover_public_actions_observations_and_independent_review(
    tmp_path: Path,
) -> None:
    acceptance = _acceptance()
    path = tmp_path / "acceptance.yaml"
    path.write_text(acceptance.model_dump_json(), encoding="utf-8")

    restored = load_acceptance(path)

    assert restored == acceptance
    assert {
        "open",
        "activate_sheet",
        "set_cell",
        "set_range",
        "recalculate",
        "execute_python_macro",
        "dispatch_control_event",
        "send_keyboard_event",
        "save",
        "close",
        "reopen",
        "export_pdf",
    } <= {item.value for item in ActionKind}
    assert {
        "cell_value",
        "cell_formula",
        "cell_type",
        "cell_error",
        "range_values",
        "sheet_state",
        "named_ranges",
        "embedded_scripts",
        "controls_bindings",
        "files_created",
        "mocked_calls",
        "screenshots",
        "runtime_errors",
    } <= {item.value for item in ObservationKind}
    with pytest.raises(ValidationError, match="must be independent"):
        MigrationMetadata(
            id="bad",
            title="Bad",
            authored_by="same",
            reviewed_by="same",
            requirements=["requirement"],
        )


def test_checked_in_yaml_models_save_close_reopen_as_first_class_actions() -> None:
    acceptance = load_acceptance(Path("examples/scenarios/public-acceptance.yaml"))

    assert [step.action.kind for step in acceptance.scenario.steps] == [
        ActionKind.OPEN,
        ActionKind.RECALCULATE,
        ActionKind.SAVE,
        ActionKind.CLOSE,
        ActionKind.REOPEN,
        ActionKind.EXPORT_PDF,
    ]
    assert acceptance.scenario.steps[4].observations_after[0].expected == ObservationValue(
        kind=ValueKind.NUMBER,
        value=42,
    )


def test_required_unavailable_and_wrong_expected_value_fail_closed() -> None:
    acceptance = _acceptance(tolerance=0.01)
    unavailable = FakeScenarioRunner(
        "fake_target",
        {"total": 10.0},
        statuses={"open": GateExecutionStatus.UNAVAILABLE},
    ).run(b"ods", acceptance.environment, acceptance.scenario)
    wrong = FakeScenarioRunner("fake_target", {"total": 10.02}).run(
        b"ods", acceptance.environment, acceptance.scenario
    )
    missing = FakeScenarioRunner("fake_target", {}).run(
        b"ods", acceptance.environment, acceptance.scenario
    )

    unavailable_evaluation = evaluate_trace(acceptance, unavailable)
    wrong_evaluation = evaluate_trace(acceptance, wrong)
    missing_evaluation = evaluate_trace(acceptance, missing)

    assert unavailable_evaluation.status is GateExecutionStatus.FAILED
    assert "action open was unavailable" in unavailable_evaluation.required_failures
    assert wrong_evaluation.status is GateExecutionStatus.FAILED
    assert wrong_evaluation.assertions[0].reason == "numeric values exceed declared tolerance"
    assert missing_evaluation.status is GateExecutionStatus.FAILED
    assert missing_evaluation.assertions[0].reason == "was not observed"


def test_fake_runner_preserves_every_action_status_deterministically() -> None:
    statuses = list(GateExecutionStatus)
    scenario = Scenario(
        id="all-statuses",
        steps=[
            ScenarioStep(
                id=f"step-{index}",
                action=Action(kind=ActionKind.OPEN, required=False),
            )
            for index, _status in enumerate(statuses)
        ],
    )
    mapping = {step.id: status for step, status in zip(scenario.steps, statuses, strict=True)}
    runner = FakeScenarioRunner("fake_target", {}, statuses=mapping)

    first = runner.run(b"same", EnvironmentManifest(), scenario)
    second = runner.run(b"same", EnvironmentManifest(), scenario)

    assert [step.status for step in first.steps] == statuses
    assert first.model_dump_json() == second.model_dump_json()
    assert first.status is GateExecutionStatus.PASSED


def test_run_writes_hashed_json_and_markdown_and_detects_tampering(tmp_path: Path) -> None:
    workbook = _write_ods(tmp_path / "target.ods")
    evidence = tmp_path / "evidence"

    manifest = run_acceptance(
        acceptance=_acceptance(),
        workbook=workbook,
        evidence_dir=evidence,
        runner=FakePathRunner({"total": 10.0}),
    )

    assert manifest.status is GateExecutionStatus.PASSED
    assert inspect_acceptance_evidence(evidence) == manifest
    assert "Cached Excel values are not used" in (evidence / "report.md").read_text()
    trace_path = evidence / "trace.json"
    trace_path.write_text(trace_path.read_text(encoding="utf-8") + " ", encoding="utf-8")
    with pytest.raises(ValueError, match="hash mismatch"):
        inspect_acceptance_evidence(evidence)

    manifest_path = evidence / "manifest.json"
    manifest_payload = json.loads(manifest_path.read_text(encoding="utf-8"))
    del manifest_payload["file_sha256"]["report.md"]
    manifest_path.write_text(json.dumps(manifest_payload), encoding="utf-8")
    with pytest.raises(ValueError, match="hashes must cover"):
        inspect_acceptance_evidence(evidence)


def test_cli_exposes_all_commands_and_diffs_typed_traces(tmp_path: Path) -> None:
    acceptance = _acceptance()
    first = FakeScenarioRunner("fake_source", {"total": 10.0}).run(
        b"ods", acceptance.environment, acceptance.scenario
    )
    second = FakeScenarioRunner("fake_target", {"total": 10.0}).run(
        b"ods", acceptance.environment, acceptance.scenario
    )
    first_path = tmp_path / "first.json"
    second_path = tmp_path / "second.json"
    first_path.write_text(first.model_dump_json(), encoding="utf-8")
    second_path.write_text(second.model_dump_json(), encoding="utf-8")
    runner = CliRunner()

    help_result = runner.invoke(cli, ["--help"])
    difference = runner.invoke(cli, ["diff", str(first_path), str(second_path)])

    assert {"run", "inspect", "diff", "mutate", "report"} <= set(help_result.output.split())
    assert difference.exit_code == 0
    assert json.loads(difference.output)["status"] == "passed"


def test_mutation_campaign_changes_only_copies_and_acceptance_kills_mutants(
    tmp_path: Path,
) -> None:
    workbook = _write_ods(tmp_path / "target.ods")
    source_hash = hashlib.sha256(workbook.read_bytes()).hexdigest()

    campaign = run_mutation_campaign(
        source_workbook=workbook,
        acceptance=_acceptance(),
        directory=tmp_path / "mutations",
        runner=FakePathRunner({"total": 99.0}),
    )
    surviving_campaign = run_mutation_campaign(
        source_workbook=workbook,
        acceptance=_acceptance(),
        directory=tmp_path / "surviving-mutations",
        runner=FakePathRunner({"total": 10.0}),
    )
    unavailable_campaign = run_mutation_campaign(
        source_workbook=workbook,
        acceptance=_acceptance(),
        directory=tmp_path / "unavailable-mutations",
        runner=FakePathRunner(
            {"total": 99.0},
            statuses={"open": GateExecutionStatus.UNAVAILABLE},
        ),
    )

    assert campaign.status is GateExecutionStatus.PASSED
    assert {case.kind for case in campaign.cases} == {"python", "formula"}
    assert {case.outcome.value for case in campaign.cases} == {"killed"}
    assert hashlib.sha256(workbook.read_bytes()).hexdigest() == source_hash
    assert (tmp_path / "mutations/mutation-report.json").is_file()
    assert (tmp_path / "mutations/mutation-report.md").is_file()
    assert surviving_campaign.status is GateExecutionStatus.FAILED
    assert {case.outcome.value for case in surviving_campaign.cases} == {"survived"}
    assert unavailable_campaign.status is GateExecutionStatus.UNAVAILABLE
    assert {case.outcome.value for case in unavailable_campaign.cases} == {"inconclusive"}


def test_report_requires_verified_passing_evidence(tmp_path: Path) -> None:
    workbook = _write_ods(tmp_path / "target.ods")
    run_acceptance(
        acceptance=_acceptance(),
        workbook=workbook,
        evidence_dir=tmp_path / "acceptance-evidence",
        runner=FakePathRunner({"total": 10.0}),
    )

    report = build_migration_report(tmp_path)

    assert report["status"] == "passed"
    assert report["evidence"][0]["migration_id"] == "invoice-migration"


def test_report_counts_killed_mutants_without_treating_expected_failures_as_regressions(
    tmp_path: Path,
) -> None:
    workbook = _write_ods(tmp_path / "target.ods")
    acceptance = _acceptance()
    run_acceptance(
        acceptance=acceptance,
        workbook=workbook,
        evidence_dir=tmp_path / "acceptance-evidence",
        runner=FakePathRunner({"total": 10.0}),
    )
    run_mutation_campaign(
        source_workbook=workbook,
        acceptance=acceptance,
        directory=tmp_path / "mutations",
        runner=FakePathRunner({"total": 99.0}),
    )

    report = build_migration_report(tmp_path)

    assert report["status"] == "passed"
    assert len(report["evidence"]) == 1
    assert report["mutation_campaigns"][0]["killed"] == 2


def test_evidence_rejects_cross_document_identity_mismatch(tmp_path: Path) -> None:
    workbook = _write_ods(tmp_path / "target.ods")
    evidence = tmp_path / "evidence"
    run_acceptance(
        acceptance=_acceptance(),
        workbook=workbook,
        evidence_dir=evidence,
        runner=FakePathRunner({"total": 10.0}),
    )
    evaluation_path = evidence / "evaluation.json"
    evaluation = json.loads(evaluation_path.read_text(encoding="utf-8"))
    evaluation["trace_id"] = "unrelated-trace"
    evaluation_path.write_text(json.dumps(evaluation), encoding="utf-8")
    manifest_path = evidence / "manifest.json"
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    manifest["file_sha256"]["evaluation.json"] = hashlib.sha256(
        evaluation_path.read_bytes()
    ).hexdigest()
    manifest_path.write_text(json.dumps(manifest), encoding="utf-8")

    with pytest.raises(ValueError, match="evaluation trace does not match"):
        inspect_acceptance_evidence(evidence)
