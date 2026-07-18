"""Real Docker-contained LibreOffice conversion integration tests."""

from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import openpyxl
import pytest

from xlsliberator.api import convert
from xlsliberator.docker_runtime import LibreOfficeDockerRuntime
from xlsliberator.libreoffice_scenario_runner import LibreOfficeScenarioRunner
from xlsliberator.scenarios.models import (
    Action,
    ActionKind,
    EnvironmentManifest,
    RuntimeTrace,
    Scenario,
    ScenarioStep,
)
from xlsliberator.validation_models import GateExecutionStatus


@pytest.mark.integration
@pytest.mark.docker
def test_convert_xlsx_to_ods_reopens_in_real_libreoffice(
    tmp_path: Path,
    skip_if_no_lo: None,
) -> None:
    """Convert a generated XLSX and verify LibreOffice recalculates the ODS."""
    input_path = tmp_path / "input.xlsx"
    output_path = tmp_path / "output.ods"
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Sheet1"
    sheet["A1"] = 2
    sheet["A2"] = 3
    sheet["A3"] = "=SUM(A1:A2)"
    sheet["B1"] = "real"
    workbook.save(input_path)
    workbook.close()

    report = convert(input_path, output_path, embed_macros=False, use_agent=False)

    assert report.success
    assert output_path.exists()

    runtime = LibreOfficeDockerRuntime()
    validation = runtime.validate_document(output_path)
    assert validation["success"] is True
    assert all(stage["status"] == "passed" for stage in validation["data"]["stages"].values())
    assert validation["data"]["source_mutated"] is False

    formula = runtime.request(
        {
            "op": "read_cell",
            "ods_path": str(output_path),
            "sheet_name": "Sheet1",
            "cell_address": "A3",
        }
    )
    text = runtime.request(
        {
            "op": "read_cell",
            "ods_path": str(output_path),
            "sheet_name": "Sheet1",
            "cell_address": "B1",
        }
    )
    assert formula["success"] is True
    assert formula["data"]["value"] == 5
    assert text["success"] is True
    assert text["data"]["value"] == "real"


@pytest.mark.integration
@pytest.mark.docker
def test_document_inspection_and_repairs_run_in_disposable_container(
    tmp_path: Path,
    skip_if_no_lo: None,
) -> None:
    """Exercise inspection and repair operations without a host UNO import."""
    input_path = tmp_path / "input.xlsx"
    output_path = tmp_path / "output.ods"
    repaired_path = tmp_path / "repaired.ods"
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Data"
    sheet["A1"] = 4
    sheet["A2"] = 6
    sheet["A3"] = "=SUM(A1:A2)"
    workbook.save(input_path)
    workbook.close()

    runtime = LibreOfficeDockerRuntime()
    runtime.convert(input_path, output_path)

    inspection = runtime.request(
        {
            "op": "inspect_document_cells",
            "ods_path": str(output_path),
            "cells": [{"sheet": "Data", "address": "A3"}],
        }
    )
    assert inspection["success"] is True
    assert inspection["data"]["cells"][0]["value"] == 10
    assert inspection["data"]["cells"][0]["formula"]

    formulas = runtime.request(
        {
            "op": "list_formula_cells",
            "ods_path": str(output_path),
        }
    )
    assert formulas["success"] is True
    assert formulas["data"]["formula_count"] == 1
    assert formulas["data"]["cells"][0]["address"] == "A3"

    repaired = runtime.request(
        {
            "op": "apply_document_repairs",
            "ods_path": str(output_path),
            "output_path": str(repaired_path),
            "formula_repairs": [],
            "named_ranges": [
                {
                    "name": "Totals",
                    "content": "$Data.$A$1:$A$3",
                    "base_sheet": 0,
                    "base_column": 0,
                    "base_row": 0,
                }
            ],
        }
    )
    assert repaired["success"] is True
    assert repaired["data"]["named_ranges_added"] == 1
    assert repaired_path.is_file()

    validation = runtime.validate_document(repaired_path)
    assert validation["success"] is True
    assert validation["data"]["source_mutated"] is False


@pytest.mark.integration
@pytest.mark.docker
def test_checked_in_ods_scenario_runs_in_disposable_libreoffice_container(
    tmp_path: Path,
    skip_if_no_lo: None,
) -> None:
    """Run the shared scenario model in Docker and prove source immutability."""
    fixture_dir = Path(__file__).parents[1] / "fixtures" / "scenarios"
    fixture = fixture_dir / "basic.ods"
    scenario = Scenario.model_validate_json(
        (fixture_dir / "basic_scenario.json").read_text(encoding="utf-8")
    )
    target = tmp_path / "target.ods"
    target.write_bytes(fixture.read_bytes())
    before = target.read_bytes()

    trace = LibreOfficeScenarioRunner().run(target, EnvironmentManifest(), scenario)

    assert trace.status is GateExecutionStatus.PASSED, trace.model_dump(mode="json")
    assert target.read_bytes() == before
    assert trace.workbook_hash_before == trace.workbook_hash_after
    assert trace.runtime_identity.runtime_kind == "libreoffice_docker"
    assert trace.runtime_identity.runtime_version == "26.2.4.2"
    assert trace.runtime_identity.image_digest
    assert trace.runtime_identity.base_image_digest
    assert trace.runtime_identity.package_manifest
    assert trace.runtime_identity.pyuno_identity["pyuno_native_sha256"]
    assert trace.runtime_identity.container_configuration["exit_code"] == 0
    assert trace.runtime_identity.metadata["profile_identifier"]
    assert trace.runtime_identity.metadata["pipe_name"]
    steps = {step.step_id: step for step in trace.steps}
    assert steps["recalculate"].observations_after["recalculated_formula"].value == 10
    assert steps["set-range"].observations_after["range_value"].value == "range"
    assert steps["reopen"].observations_after["reopened_formula"].value == 10
    assert any(item.startswith("exported output:") for item in steps["export"].evidence)
    assert all(step.status is GateExecutionStatus.PASSED for step in trace.steps)


@pytest.mark.integration
@pytest.mark.docker
def test_target_formula_parser_round_trips_in_target_document_context(
    tmp_path: Path,
    skip_if_no_lo: None,
) -> None:
    """Use FormulaParser inside Docker with the target sheet/cell context."""
    fixture = Path(__file__).parents[1] / "fixtures" / "scenarios" / "basic.ods"
    target = tmp_path / "target.ods"
    target.write_bytes(fixture.read_bytes())
    before = target.read_bytes()
    runtime = LibreOfficeDockerRuntime()
    identity = runtime.resolve_identity()

    parsed = runtime.parse_formula(
        target,
        "=SUM(A1:A2)",
        sheet_name="Sheet1",
        cell_address="A3",
        image_id=identity.image_id,
    )

    assert parsed["success"] is True, parsed
    assert parsed["data"]["document_context"] == "target"
    assert parsed["data"]["tokens"]
    assert parsed["data"]["roundtrip_formula"]
    assert parsed["data"]["roundtrip_equivalent"] is True
    assert parsed["data"]["parser_accepted"] is True
    assert parsed["data"]["syntax_errors"] == []
    assert target.read_bytes() == before


@pytest.mark.integration
@pytest.mark.docker
def test_formula_repair_is_applied_transactionally_inside_docker(
    tmp_path: Path,
    skip_if_no_lo: None,
) -> None:
    """Write a candidate ODS in Docker while preserving the original package."""
    fixture = Path(__file__).parents[1] / "fixtures" / "scenarios" / "basic.ods"
    source = tmp_path / "source.ods"
    candidate = tmp_path / "candidate.ods"
    source.write_bytes(fixture.read_bytes())
    before = source.read_bytes()
    runtime = LibreOfficeDockerRuntime()

    repaired = runtime.request(
        {
            "op": "apply_document_repairs",
            "ods_path": str(source),
            "output_path": str(candidate),
            "formula_repairs": [
                {
                    "sheet": "Sheet1",
                    "address": "A3",
                    "formula": "=SUM(A1:A2)",
                    "rule_name": "integration_fixture",
                }
            ],
            "named_ranges": [],
        }
    )
    observed = runtime.request(
        {
            "op": "read_cell",
            "ods_path": str(candidate),
            "sheet_name": "Sheet1",
            "cell_address": "A3",
        }
    )

    assert repaired["success"] is True, repaired
    assert repaired["data"]["formulas_applied"] == 1
    assert candidate.is_file()
    assert source.read_bytes() == before
    assert candidate.read_bytes() != before
    assert observed["success"] is True, observed
    assert "SUM" in observed["data"]["formula"]


@pytest.mark.integration
@pytest.mark.docker
def test_concurrent_scenarios_use_independent_containers_profiles_and_pipes(
    tmp_path: Path,
    skip_if_no_lo: None,
) -> None:
    """Prove concurrent jobs never share an office process or execution workspace."""
    fixture_dir = Path(__file__).parents[1] / "fixtures" / "scenarios"
    fixture = fixture_dir / "basic.ods"
    scenario = Scenario.model_validate_json(
        (fixture_dir / "basic_scenario.json").read_text(encoding="utf-8")
    )
    workbooks = [tmp_path / f"target-{index}.ods" for index in range(2)]
    for workbook in workbooks:
        workbook.write_bytes(fixture.read_bytes())

    def execute(workbook: Path) -> RuntimeTrace:
        return LibreOfficeScenarioRunner().run(workbook, EnvironmentManifest(), scenario)

    with ThreadPoolExecutor(max_workers=2) as executor:
        traces = list(executor.map(execute, workbooks))

    assert all(trace.status is GateExecutionStatus.PASSED for trace in traces)
    container_names = {
        trace.runtime_identity.container_configuration["container_name"] for trace in traces
    }
    job_ids = {trace.runtime_identity.container_configuration["job_id"] for trace in traces}
    profiles = {trace.runtime_identity.metadata["profile_identifier"] for trace in traces}
    pipes = {trace.runtime_identity.metadata["pipe_name"] for trace in traces}
    assert len(container_names) == len(job_ids) == len(profiles) == len(pipes) == 2


@pytest.mark.integration
@pytest.mark.docker
@pytest.mark.parametrize("kind", [ActionKind.INVOKE_MACRO, ActionKind.CLICK_CONTROL])
def test_macro_and_control_actions_are_unavailable_without_explicit_capability(
    tmp_path: Path,
    skip_if_no_lo: None,
    kind: ActionKind,
) -> None:
    """A missing macro grant must never be treated as successful execution."""
    fixture = Path(__file__).parents[1] / "fixtures" / "scenarios" / "basic.ods"
    target = tmp_path / "target.ods"
    target.write_bytes(fixture.read_bytes())
    scenario = Scenario(
        id=f"missing-capability-{kind.value}",
        steps=[
            ScenarioStep(id="open", action=Action(kind=ActionKind.OPEN)),
            ScenarioStep(
                id="protected-action",
                action=Action(
                    kind=kind,
                    parameters={
                        "script_uri": "vnd.sun.star.script:Missing.py$Main?language=Python&location=document",
                        "control_name": "Missing",
                    },
                ),
            ),
        ],
    )

    trace = LibreOfficeScenarioRunner().run(target, EnvironmentManifest(), scenario)

    assert trace.status is GateExecutionStatus.UNAVAILABLE
    protected = next(step for step in trace.steps if step.step_id == "protected-action")
    assert protected.status is GateExecutionStatus.UNAVAILABLE
    assert protected.error is not None
    assert "macro_execution" in protected.error["message"]
