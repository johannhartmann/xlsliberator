"""Scenario schema, normalization, trace diff, and evidence tests."""

import hashlib
import json
from datetime import date
from pathlib import Path

from click.testing import CliRunner

from xlsliberator.cli import cli
from xlsliberator.ir_models import WorkbookIR
from xlsliberator.scenarios.diff import diff_traces
from xlsliberator.scenarios.evidence import EvidenceBundleWriter, inspect_evidence_bundle
from xlsliberator.scenarios.models import (
    Action,
    ActionKind,
    ComparisonRules,
    EnvironmentManifest,
    ObservationKind,
    ObservationRequest,
    RuntimeTrace,
    Scenario,
    ScenarioStep,
    ValueKind,
)
from xlsliberator.scenarios.normalize import normalize_value
from xlsliberator.scenarios.runner import FakeScenarioRunner
from xlsliberator.validation_models import (
    GateExecutionStatus,
    InventoryDiff,
    WorkbookArtifactIR,
)


def _scenario(*, tolerance: float = 0.0) -> Scenario:
    return Scenario(
        id="calculation",
        steps=[
            ScenarioStep(
                id="open",
                action=Action(kind=ActionKind.OPEN),
                observations_after=[
                    ObservationRequest(
                        id="value",
                        kind=ObservationKind.CELL,
                        selector={"sheet": "Sheet1", "address": "A1"},
                        comparison=ComparisonRules(absolute_tolerance=tolerance),
                    ),
                    ObservationRequest(id="empty", kind=ObservationKind.CELL),
                ],
            )
        ],
    )


def test_scenario_json_round_trip_is_stable() -> None:
    scenario = _scenario()

    restored = Scenario.model_validate_json(scenario.model_dump_json())

    assert restored == scenario
    assert restored.schema_version == "1.0.0"


def test_normalization_preserves_boolean_empty_error_date_and_whitespace() -> None:
    assert normalize_value(True).kind is ValueKind.BOOLEAN
    assert normalize_value(1).kind is ValueKind.NUMBER
    assert normalize_value(None).kind is ValueKind.EMPTY_CELL
    assert normalize_value("").kind is ValueKind.EMPTY_STRING
    assert normalize_value("#REF!").error_type == "#REF!"
    assert normalize_value("#FUTURE!", error_type="#FUTURE!").error_type == "#FUTURE!"
    assert normalize_value("  text  ").value == "  text  "


def test_diff_compares_error_and_date_semantics_not_runtime_metadata() -> None:
    scenario = _scenario()
    environment = EnvironmentManifest()
    source = FakeScenarioRunner(
        "fake_source",
        {
            "value": normalize_value("#DIV/0!", error_type="#DIV/0!", cell_type="ExcelError"),
            "empty": normalize_value(date(2026, 7, 12), date_system="1900"),
        },
    ).run(b"x", environment, scenario)
    target = FakeScenarioRunner(
        "fake_target",
        {
            "value": normalize_value("#DIV/0!", error_type="#DIV/0!", cell_type="FORMULA"),
            "empty": normalize_value(date(2026, 7, 12), date_system="1900"),
        },
    ).run(b"x", environment, scenario)

    result = diff_traces(source, target, scenario)

    assert result.status is GateExecutionStatus.PASSED


def test_fake_source_and_target_consume_same_scenario_and_diff_deterministically() -> None:
    scenario = _scenario(tolerance=0.01)
    environment = EnvironmentManifest()
    workbook = b"fake workbook"
    source = FakeScenarioRunner("fake_source", {"value": 10.0, "empty": None}).run(
        workbook, environment, scenario
    )
    target = FakeScenarioRunner("fake_target", {"value": 10.005, "empty": None}).run(
        workbook, environment, scenario
    )

    result = diff_traces(source, target, scenario)

    assert result.status is GateExecutionStatus.PASSED
    assert result.equivalent
    assert all(item.matched for item in result.differences)


def test_diff_does_not_equate_boolean_number_or_empty_string_cell() -> None:
    scenario = _scenario()
    environment = EnvironmentManifest()
    source = FakeScenarioRunner("fake_source", {"value": True, "empty": None}).run(
        b"x", environment, scenario
    )
    target = FakeScenarioRunner("fake_target", {"value": 1, "empty": ""}).run(
        b"x", environment, scenario
    )

    result = diff_traces(source, target, scenario)

    assert result.status is GateExecutionStatus.FAILED
    reasons = {item.observation_id: item.reason for item in result.differences}
    assert reasons["value"].startswith("type differs")
    assert reasons["empty"] == "empty string and empty cell are distinct"


def test_evidence_bundle_records_hashes_traces_and_runtime_identity(tmp_path: Path) -> None:
    source_file = tmp_path / "source.xlsx"
    output_file = tmp_path / "output.ods"
    source_file.write_bytes(b"source")
    output_file.write_bytes(b"output")
    scenario = _scenario()
    environment = EnvironmentManifest(
        declared_capabilities=["macro_execution"],
        granted_capabilities=["macro_execution"],
    )
    source = FakeScenarioRunner("fake_source", {"value": 1, "empty": None}).run(
        source_file.read_bytes(), environment, scenario
    )
    target = FakeScenarioRunner("fake_target", {"value": 1, "empty": None}).run(
        output_file.read_bytes(), environment, scenario
    )
    difference = diff_traces(source, target, scenario)
    source_inventory = WorkbookArtifactIR(
        workbook=WorkbookIR(file_path=str(source_file), file_format="xlsx")
    )
    target_inventory = WorkbookArtifactIR(
        inventory_role="target",
        workbook=WorkbookIR(file_path=str(output_file), file_format="ods"),
    )
    inventory_difference = InventoryDiff(
        source_inventory_sha256="a" * 64,
        target_inventory_sha256="b" * 64,
    )
    bundle = tmp_path / "bundle"

    manifest = EvidenceBundleWriter(bundle).write(
        source_workbook=source_file,
        output=output_file,
        environment=environment,
        scenario=scenario,
        source_trace=source,
        target_traces={"libreoffice": target},
        diffs=[difference],
        source_inventory=source_inventory,
        target_inventories={"libreoffice": target_inventory},
        inventory_diffs=[inventory_difference],
    )

    assert manifest.source_workbook_hash == hashlib.sha256(b"source").hexdigest()
    assert manifest.output_hash == hashlib.sha256(b"output").hexdigest()
    assert manifest.granted_capabilities == ["macro_execution"]
    assert manifest.source_inventory == "source-inventory.json"
    assert manifest.target_inventories == {"libreoffice": "target-libreoffice-inventory.json"}
    assert manifest.inventory_diffs == ["inventory-diff-1.json"]
    assert inspect_evidence_bundle(bundle) == manifest


def test_cli_validates_scenario_diffs_traces_and_inspects_evidence(tmp_path: Path) -> None:
    runner = CliRunner()
    scenario = _scenario()
    environment = EnvironmentManifest()
    source = FakeScenarioRunner("fake_source", {"value": 1, "empty": None}).run(
        b"x", environment, scenario
    )
    target = FakeScenarioRunner("fake_target", {"value": 1, "empty": None}).run(
        b"x", environment, scenario
    )
    scenario_path = tmp_path / "scenario.json"
    source_path = tmp_path / "source.json"
    target_path = tmp_path / "target.json"
    scenario_path.write_text(scenario.model_dump_json())
    source_path.write_text(source.model_dump_json())
    target_path.write_text(target.model_dump_json())

    validated = runner.invoke(cli, ["scenario-validate", str(scenario_path)])
    compared = runner.invoke(
        cli, ["trace-diff", str(scenario_path), str(source_path), str(target_path)]
    )

    assert validated.exit_code == 0
    assert json.loads(validated.output)["id"] == scenario.id
    assert compared.exit_code == 0
    assert json.loads(compared.output)["status"] == "passed"


def test_runtime_trace_schema_rejects_unknown_fields() -> None:
    scenario = _scenario()
    trace = FakeScenarioRunner("fake_source", {"value": 1, "empty": None}).run(
        b"x", EnvironmentManifest(), scenario
    )
    payload = trace.model_dump(mode="json")
    payload["invented_success"] = True

    try:
        RuntimeTrace.model_validate(payload)
    except ValueError as exc:
        assert "invented_success" in str(exc)
    else:
        raise AssertionError("unknown trace fields must be rejected")
