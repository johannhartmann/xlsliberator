"""Tests for the Docker-only LibreOffice scenario runner boundary."""

from __future__ import annotations

import hashlib
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

from click.testing import CliRunner

from xlsliberator.cli import cli
from xlsliberator.docker_runtime import DockerRuntimeIdentity, DockerRuntimeUnavailable
from xlsliberator.libreoffice_scenario_runner import LibreOfficeScenarioRunner
from xlsliberator.scenarios.models import (
    Action,
    ActionKind,
    EnvironmentManifest,
    ObservationKind,
    ObservationRequest,
    RuntimeIdentity,
    RuntimeTrace,
    Scenario,
    ScenarioStep,
)
from xlsliberator.validation_models import GateExecutionStatus


def _scenario() -> Scenario:
    return Scenario(
        id="target-smoke",
        steps=[
            ScenarioStep(
                id="open",
                action=Action(kind=ActionKind.OPEN),
                observations_after=[
                    ObservationRequest(
                        id="a1",
                        kind=ObservationKind.CELL,
                        selector={"sheet": "Sheet1", "address": "A1"},
                    )
                ],
            )
        ],
    )


class FakeDockerRuntime:
    def resolve_identity(self, *, probe: bool = True) -> DockerRuntimeIdentity:
        assert probe
        return DockerRuntimeIdentity(
            image_reference="xlsliberator-libreoffice:26.2.4.2",
            image_id="sha256:fixed",
            version="26.2.4.2",
            architecture="arm64",
            probe={
                "office_executable": "/opt/libreoffice26.2/program/soffice",
                "office_sha256": "a" * 64,
                "base_image_digest": "sha256:base",
                "python_version": "3.12.13",
                "uno_module": "/opt/libreoffice26.2/program/uno.py",
                "uno_module_sha256": "b" * 64,
                "pyuno_native_module": "/opt/libreoffice26.2/program/pyuno.so",
                "pyuno_native_sha256": "c" * 64,
                "worker_wrapper": "/usr/local/bin/runtime-entrypoint",
                "worker_wrapper_sha256": "d" * 64,
                "installed_package_manifest": [{"name": "libobasis26.2-pyuno"}],
            },
        )

    def request(self, payload: dict[str, Any], *, _identity: str | None = None) -> dict[str, Any]:
        assert payload["op"] == "run_scenario"
        assert _identity == "sha256:fixed"
        now = datetime.now(UTC).isoformat()
        return {
            "success": True,
            "data": {
                "scenario_id": "target-smoke",
                "status": "passed",
                "started_at": now,
                "ended_at": now,
                "steps": [
                    {
                        "step_id": "open",
                        "action": "open",
                        "status": "passed",
                        "started_at": now,
                        "ended_at": now,
                        "observations_after": {
                            "a1": {"kind": "number", "value": 42, "cell_type": "VALUE"}
                        },
                    }
                ],
                "runtime": {
                    "profile_identifier": "unique-profile",
                    "pipe_name": "unique-pipe",
                    "office_exit_code": 0,
                },
                "resource_policy": {"network": "none"},
                "container_exit_code": 0,
                "container_name": "unique-container",
                "job_id": "unique-job",
                "logs": ["office log"],
                "final_working_copy_sha256": "e" * 64,
                "source_mutated": False,
            },
        }


def test_target_runner_maps_complete_identity_and_preserves_source(tmp_path: Path) -> None:
    workbook = tmp_path / "target.ods"
    workbook.write_bytes(b"fake ODS")

    trace = LibreOfficeScenarioRunner(runtime=FakeDockerRuntime()).run(
        workbook,
        EnvironmentManifest(),
        _scenario(),
    )

    assert trace.status is GateExecutionStatus.PASSED
    assert trace.workbook_hash_before == hashlib.sha256(b"fake ODS").hexdigest()
    assert trace.workbook_hash_after == trace.workbook_hash_before
    assert trace.runtime_identity.image_digest == "sha256:fixed"
    assert trace.runtime_identity.container_configuration["container_name"] == "unique-container"
    assert trace.runtime_identity.metadata["pipe_name"] == "unique-pipe"
    assert trace.steps[0].observations_after["a1"].value == 42
    assert trace.logs == ["office log"]


def test_missing_docker_is_explicitly_unavailable_without_fallback(tmp_path: Path) -> None:
    class UnavailableRuntime:
        def resolve_identity(self, *, probe: bool = True) -> DockerRuntimeIdentity:
            raise DockerRuntimeUnavailable("Docker missing; host fallback is disabled")

        def request(
            self, payload: dict[str, Any], *, _identity: str | None = None
        ) -> dict[str, Any]:
            raise AssertionError("request must not run when identity resolution fails")

    workbook = tmp_path / "target.ods"
    workbook.write_bytes(b"fake ODS")

    trace = LibreOfficeScenarioRunner(runtime=UnavailableRuntime()).run(
        workbook,
        EnvironmentManifest(),
        _scenario(),
    )

    assert trace.status is GateExecutionStatus.UNAVAILABLE
    assert trace.error is not None
    assert "host fallback is disabled" in trace.error["message"]


def test_scenario_run_target_cli_serializes_trace(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    workbook = tmp_path / "target.ods"
    workbook.write_bytes(b"fake ODS")
    scenario = _scenario()
    scenario_path = tmp_path / "scenario.json"
    scenario_path.write_text(scenario.model_dump_json(), encoding="utf-8")
    now = datetime.now(UTC)
    digest = hashlib.sha256(b"fake ODS").hexdigest()
    trace = RuntimeTrace(
        trace_id="trace",
        scenario_id=scenario.id,
        runtime_role="target",
        runtime_identity=RuntimeIdentity(
            runtime_kind="libreoffice_docker", runtime_version="26.2.4.2"
        ),
        environment=EnvironmentManifest(),
        status=GateExecutionStatus.PASSED,
        started_at=now,
        ended_at=now,
        workbook_hash_before=digest,
        workbook_hash_after=digest,
    )
    monkeypatch.setattr(
        "xlsliberator.libreoffice_scenario_runner.LibreOfficeScenarioRunner.run",
        lambda *_args, **_kwargs: trace,
    )
    output = tmp_path / "trace.json"

    result = CliRunner().invoke(
        cli,
        ["scenario-run-target", str(workbook), str(scenario_path), "--output", str(output)],
    )

    assert result.exit_code == 0
    assert RuntimeTrace.model_validate_json(output.read_text(encoding="utf-8")) == trace
