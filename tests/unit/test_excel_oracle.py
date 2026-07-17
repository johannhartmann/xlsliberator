"""Tests for the Excel source-oracle protocol without executing Excel."""

from pathlib import Path
from typing import Any

from xlsliberator.excel_oracle import (
    FakeExcelOracle,
    OracleRunResult,
    OracleTransportUnavailable,
    UnavailableExcelOracle,
    WindowsExcelOracleClient,
    load_source_trace_fixture,
)
from xlsliberator.execution_sandbox import SandboxBackendKind, SandboxPolicy
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
from xlsliberator.windows_excel_worker import execute_request


class _Transport:
    def __init__(
        self, response: dict[str, Any] | None = None, error: Exception | None = None
    ) -> None:
        self.response = response or {}
        self.error = error
        self.requests: list[dict[str, Any]] = []

    def submit(self, request: dict[str, Any], timeout_seconds: float) -> dict[str, Any]:
        del timeout_seconds
        self.requests.append(request)
        if self.error:
            raise self.error
        return self.response


def _all_actions_scenario() -> Scenario:
    return Scenario(
        id="all-actions",
        steps=[
            ScenarioStep(
                id=kind.value,
                action=Action(
                    kind=kind,
                    required=kind
                    not in {
                        ActionKind.CLICK_CONTROL,
                        ActionKind.REFRESH_DATA,
                        ActionKind.PRINT,
                        ActionKind.EXPORT,
                    },
                ),
                observations_after=[
                    ObservationRequest(id=f"observation-{kind.value}", kind=ObservationKind.CELL)
                ],
            )
            for kind in ActionKind
        ],
    )


def test_fake_oracle_covers_every_declared_action(tmp_path: Path) -> None:
    workbook = tmp_path / "book.xlsm"
    workbook.write_bytes(b"workbook")
    scenario = _all_actions_scenario()

    result = FakeExcelOracle().run(workbook, EnvironmentManifest(), scenario)

    assert result.succeeded
    assert result.trace is not None
    assert {step.action for step in result.trace.steps} == set(ActionKind)
    assert result.trace.runtime_identity.runtime_kind == "fake_excel_oracle"


def test_fake_required_failure_fails_trace(tmp_path: Path) -> None:
    workbook = tmp_path / "book.xlsx"
    workbook.write_bytes(b"workbook")

    result = FakeExcelOracle(action_statuses={"recalculate": GateExecutionStatus.FAILED}).run(
        workbook, EnvironmentManifest(), _all_actions_scenario()
    )

    assert result.status is GateExecutionStatus.FAILED
    assert result.trace is not None
    assert not result.succeeded


def test_unconfigured_or_unreachable_oracle_is_explicitly_unavailable(tmp_path: Path) -> None:
    workbook = tmp_path / "book.xlsx"
    workbook.write_bytes(b"workbook")
    scenario = _all_actions_scenario()
    environment = EnvironmentManifest()

    local = UnavailableExcelOracle().run(workbook, environment, scenario)
    remote = WindowsExcelOracleClient(_Transport(error=OracleTransportUnavailable("offline"))).run(
        workbook, environment, scenario
    )

    assert local.status is GateExecutionStatus.UNAVAILABLE
    assert remote.status is GateExecutionStatus.UNAVAILABLE
    assert local.trace is None and remote.trace is None


def test_client_submits_versioned_payload_and_accepts_real_source_trace(tmp_path: Path) -> None:
    workbook = tmp_path / "book.xlsx"
    workbook.write_bytes(b"workbook")
    scenario = _all_actions_scenario()
    environment = EnvironmentManifest()
    fake_trace = FakeExcelOracle().run(workbook, environment, scenario).trace
    assert fake_trace is not None
    payload = fake_trace.model_dump(mode="json")
    payload["runtime_role"] = "source"
    payload["runtime_identity"] = RuntimeIdentity(
        runtime_kind="microsoft_excel",
        runtime_version="16.0.19029",
        container_configuration={
            "sandbox_policy": SandboxPolicy(backend=SandboxBackendKind.REMOTE_WORKER).model_dump(
                mode="json"
            )
        },
    ).model_dump(mode="json")
    transport = _Transport(
        OracleRunResult(
            status=GateExecutionStatus.PASSED,
            trace=RuntimeTrace.model_validate(payload),
        ).model_dump(mode="json")
    )

    result = WindowsExcelOracleClient(transport).run(workbook, environment, scenario)

    assert result.succeeded
    request = transport.requests[0]
    assert request["schema_version"] == "1.0.0"
    assert request["workbook_name"] == "book.xlsx"
    assert request["workbook_base64"]
    assert request["execution_kind"] == "source_oracle"
    assert request["sandbox_policy"]["backend"] == "remote_worker"


def test_malformed_or_non_source_response_fails_closed(tmp_path: Path) -> None:
    workbook = tmp_path / "book.xlsx"
    workbook.write_bytes(b"workbook")
    scenario = _all_actions_scenario()

    malformed = WindowsExcelOracleClient(_Transport({"status": "passed"})).run(
        workbook, EnvironmentManifest(), scenario
    )
    fake_target = FakeExcelOracle().run(workbook, EnvironmentManifest(), scenario).trace
    assert fake_target is not None
    target_payload = fake_target.model_dump(mode="json")
    target_payload["runtime_role"] = "target"
    wrong_role = WindowsExcelOracleClient(
        _Transport(
            OracleRunResult(
                status=GateExecutionStatus.PASSED,
                trace=RuntimeTrace.model_validate(target_payload),
            ).model_dump(mode="json")
        )
    ).run(workbook, EnvironmentManifest(), scenario)

    assert malformed.status is GateExecutionStatus.FAILED
    assert wrong_role.status is GateExecutionStatus.FAILED


def test_source_trace_fixture_requires_real_excel_identity(tmp_path: Path) -> None:
    workbook = tmp_path / "book.xlsx"
    workbook.write_bytes(b"workbook")
    scenario = _all_actions_scenario()
    trace = FakeExcelOracle().run(workbook, EnvironmentManifest(), scenario).trace
    assert trace is not None
    fixture = tmp_path / "trace.json"
    fixture.write_text(trace.model_dump_json())

    try:
        load_source_trace_fixture(fixture)
    except ValueError as exc:
        assert "Microsoft Excel" in str(exc)
    else:
        raise AssertionError("fake source traces must not masquerade as Excel fixtures")


def test_windows_worker_is_unavailable_off_windows_without_importing_com() -> None:
    result = execute_request(
        {
            "schema_version": "1.0.0",
            "environment": EnvironmentManifest().model_dump(mode="json"),
            "scenario": _all_actions_scenario().model_dump(mode="json"),
            "workbook_name": "book.xlsx",
            "workbook_base64": "",
        }
    )

    assert result.status is GateExecutionStatus.UNAVAILABLE
    assert result.trace is None
