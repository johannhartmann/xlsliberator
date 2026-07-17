"""Windows-only, one-job Microsoft Excel source-oracle worker.

The module imports COM dependencies only inside ``execute_request``. Importing
it on Linux/macOS never discovers or launches an office application.
"""

from __future__ import annotations

import base64
import contextlib
import hashlib
import importlib
import json
import platform
import sys
import tempfile
import zipfile
from datetime import UTC, date, datetime
from pathlib import Path
from typing import Any
from uuid import uuid4

from xlsliberator.excel_oracle import ORACLE_PROTOCOL_VERSION, OracleRunResult
from xlsliberator.execution_sandbox import (
    ExecutionKind,
    SandboxBackendKind,
    SandboxPolicy,
)
from xlsliberator.scenarios.models import (
    EnvironmentManifest,
    ObservationKind,
    ObservationRequest,
    ObservationValue,
    RuntimeIdentity,
    RuntimeTrace,
    Scenario,
    StepResult,
)
from xlsliberator.scenarios.normalize import normalize_value
from xlsliberator.validation_models import GateExecutionStatus


def execute_request(request: dict[str, Any]) -> OracleRunResult:
    """Execute one request in a fresh Excel COM process on Windows."""
    if request.get("schema_version") != ORACLE_PROTOCOL_VERSION:
        return _error(GateExecutionStatus.FAILED, "protocol_version", "unsupported schema")
    if platform.system() != "Windows":
        return _error(
            GateExecutionStatus.UNAVAILABLE,
            "windows_unavailable",
            "Microsoft Excel source execution requires the configured Windows worker",
        )
    try:
        pythoncom = importlib.import_module("pythoncom")
        win32com_client = importlib.import_module("win32com.client")
        win32process = importlib.import_module("win32process")
    except ImportError as exc:
        return _error(
            GateExecutionStatus.UNAVAILABLE,
            "pywin32_unavailable",
            f"Windows worker is missing pywin32: {exc}",
        )

    try:
        environment = EnvironmentManifest.model_validate(request["environment"])
        scenario = Scenario.model_validate(request["scenario"])
        sandbox_policy = SandboxPolicy.model_validate(request["sandbox_policy"])
        if request.get("execution_kind") != ExecutionKind.SOURCE_ORACLE.value:
            raise ValueError("source oracle request has the wrong execution kind")
        if sandbox_policy.backend not in {
            SandboxBackendKind.REMOTE_WORKER,
            SandboxBackendKind.MICROVM,
        }:
            raise ValueError("source oracle requires a remote-worker or microVM sandbox")
        workbook_bytes = base64.b64decode(request["workbook_base64"], validate=True)
        workbook_name = Path(str(request["workbook_name"])).name
    except Exception as exc:
        return _error(GateExecutionStatus.FAILED, "invalid_request", str(exc))

    pythoncom.CoInitialize()
    excel = None
    workbook = None
    started = datetime.now(UTC)
    before_hash = hashlib.sha256(workbook_bytes).hexdigest()
    try:
        with tempfile.TemporaryDirectory(prefix="xlsliberator-excel-oracle-") as directory:
            job_dir = Path(directory)
            source_path = job_dir / workbook_name
            source_path.write_bytes(workbook_bytes)
            current_path = source_path
            excel = win32com_client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.AskToUpdateLinks = False
            excel.EnableEvents = True
            # msoAutomationSecurityForceDisable unless macro execution was explicitly granted.
            macro_granted = "macro_execution" in environment.all_granted_capabilities
            excel.AutomationSecurity = 1 if macro_granted else 3
            hwnd = int(excel.Hwnd)
            _thread_id, excel_pid = win32process.GetWindowThreadProcessId(hwnd)
            pid_file = request.get("excel_pid_file")
            if pid_file:
                Path(str(pid_file)).write_text(str(excel_pid), encoding="ascii")
            workbook = excel.Workbooks.Open(
                str(current_path),
                UpdateLinks=0,
                ReadOnly=False,
                IgnoreReadOnlyRecommended=True,
            )
            _configure_calculation(excel, environment)
            steps: list[StepResult] = []
            for definition in scenario.steps:
                step_started = datetime.now(UTC)
                before, before_status, before_error = _capture_requests(
                    workbook, definition.observations_before, environment
                )
                action_status, action_error, workbook, current_path = _execute_action(
                    excel,
                    workbook,
                    current_path,
                    job_dir,
                    definition.action.kind.value,
                    definition.action.parameters,
                    macro_granted,
                )
                after, after_status, after_error = _capture_requests(
                    workbook, definition.observations_after, environment
                )
                statuses = [before_status, action_status, after_status]
                status = next(
                    (item for item in statuses if item is not GateExecutionStatus.PASSED),
                    GateExecutionStatus.PASSED,
                )
                error = action_error or before_error or after_error
                steps.append(
                    StepResult(
                        step_id=definition.id,
                        action=definition.action.kind,
                        status=status,
                        started_at=step_started,
                        ended_at=datetime.now(UTC),
                        observations_before=before,
                        observations_after=after,
                        error=error,
                    )
                )
                if definition.action.required and status is not GateExecutionStatus.PASSED:
                    break
            final_hash = _hash_path(current_path) if current_path.is_file() else None
            required_failed = any(
                result.status is not GateExecutionStatus.PASSED
                and next(
                    step for step in scenario.steps if step.id == result.step_id
                ).action.required
                for result in steps
            )
            identity = RuntimeIdentity(
                runtime_kind="microsoft_excel",
                runtime_version=f"{excel.Version}.{excel.Build}",
                executable_path=str(excel.Path),
                container_configuration={"sandbox_policy": sandbox_policy.model_dump(mode="json")},
                metadata={
                    "windows_version": platform.platform(),
                    "locale": environment.locale,
                    "timezone": environment.timezone,
                    "date_system": environment.date_system,
                    "calculation": int(excel.Calculation),
                    "automation_security": int(excel.AutomationSecurity),
                    "add_ins": [str(item.Name) for item in excel.AddIns if bool(item.Installed)],
                    "excel_pid": excel_pid,
                },
            )
            trace_status = (
                GateExecutionStatus.FAILED if required_failed else GateExecutionStatus.PASSED
            )
            trace = RuntimeTrace(
                trace_id=f"excel-{uuid4().hex}",
                scenario_id=scenario.id,
                runtime_role="source",
                runtime_identity=identity,
                environment=environment,
                status=trace_status,
                started_at=started,
                ended_at=datetime.now(UTC),
                workbook_hash_before=before_hash,
                workbook_hash_after=final_hash,
                steps=steps,
            )
            return OracleRunResult(status=trace_status, trace=trace)
    except Exception as exc:
        return _error(GateExecutionStatus.FAILED, "excel_execution_failed", str(exc))
    finally:
        if workbook is not None:
            with contextlib.suppress(Exception):
                workbook.Close(SaveChanges=False)
        if excel is not None:
            with contextlib.suppress(Exception):
                excel.Quit()
        pythoncom.CoUninitialize()


def _execute_action(
    excel: Any,
    workbook: Any,
    current_path: Path,
    job_dir: Path,
    action: str,
    parameters: dict[str, Any],
    macro_granted: bool,
) -> tuple[GateExecutionStatus, dict[str, Any] | None, Any, Path]:
    try:
        if action == "open":
            return GateExecutionStatus.PASSED, None, workbook, current_path
        if action == "close":
            workbook.Close(SaveChanges=False)
            return GateExecutionStatus.PASSED, None, None, current_path
        if workbook is None:
            return (
                GateExecutionStatus.FAILED,
                {"type": "workbook_closed", "message": "action requires an open workbook"},
                workbook,
                current_path,
            )
        if action == "set_cell":
            cell = workbook.Worksheets(parameters["sheet"]).Range(parameters["address"])
            if "formula" in parameters:
                cell.Formula = parameters["formula"]
            else:
                cell.Value2 = parameters.get("value")
        elif action == "set_range":
            workbook.Worksheets(parameters["sheet"]).Range(
                parameters["address"]
            ).Value2 = parameters["values"]
        elif action == "recalculate":
            excel.CalculateFullRebuild()
        elif action == "invoke_macro":
            if not macro_granted:
                return (
                    GateExecutionStatus.UNAVAILABLE,
                    {"type": "capability_denied", "message": "macro_execution was not granted"},
                    workbook,
                    current_path,
                )
            excel.Run(parameters["procedure"], *parameters.get("arguments", []))
        elif action == "activate_sheet":
            workbook.Worksheets(parameters["sheet"]).Activate()
        elif action == "save":
            workbook.Save()
        elif action == "save_as":
            destination = job_dir / Path(str(parameters.get("name", "oracle-output.xlsx"))).name
            workbook.SaveAs(str(destination))
            current_path = destination
        elif action == "reopen":
            workbook.Close(SaveChanges=True)
            workbook = excel.Workbooks.Open(str(current_path), UpdateLinks=0, ReadOnly=False)
        elif action in {"click_control", "refresh_data", "print", "export"}:
            return (
                GateExecutionStatus.UNAVAILABLE,
                {"type": "action_unavailable", "message": action},
                workbook,
                current_path,
            )
        else:
            raise ValueError(f"unsupported action: {action}")
        return GateExecutionStatus.PASSED, None, workbook, current_path
    except Exception as exc:
        return (
            GateExecutionStatus.FAILED,
            {"type": "action_failed", "message": str(exc)},
            workbook,
            current_path,
        )


def _capture_requests(
    workbook: Any,
    requests: list[ObservationRequest],
    environment: EnvironmentManifest,
) -> tuple[dict[str, ObservationValue], GateExecutionStatus, dict[str, Any] | None]:
    values: dict[str, ObservationValue] = {}
    if workbook is None and requests:
        return (
            values,
            GateExecutionStatus.FAILED,
            {
                "type": "workbook_closed",
                "message": "observation requires an open workbook",
            },
        )
    for request in requests:
        try:
            values[request.id] = _capture_observation(workbook, request, environment)
        except PermissionError as exc:
            if request.required:
                return (
                    values,
                    GateExecutionStatus.UNAVAILABLE,
                    {
                        "type": "observation_unavailable",
                        "message": str(exc),
                    },
                )
        except Exception as exc:
            if request.required:
                return (
                    values,
                    GateExecutionStatus.FAILED,
                    {
                        "type": "observation_failed",
                        "message": str(exc),
                    },
                )
    return values, GateExecutionStatus.PASSED, None


def _configure_calculation(excel: Any, environment: EnvironmentManifest) -> None:
    """Apply the scenario's explicit Excel calculation semantics."""
    modes = {
        "automatic": -4105,
        "manual": -4135,
        "automatic_except_tables": 2,
    }
    excel.Calculation = modes[environment.calculation_mode]
    excel.Iteration = environment.iterative_calculation
    excel.MaxIterations = environment.max_iterations
    excel.MaxChange = environment.max_change


def _capture_observation(
    workbook: Any, request: ObservationRequest, environment: EnvironmentManifest
) -> ObservationValue:
    if request.kind is ObservationKind.CELL:
        cell = workbook.Worksheets(request.selector["sheet"]).Range(request.selector["address"])
        value2 = cell.Value2
        typed_value = cell.Value
        value = typed_value if isinstance(typed_value, (date, datetime)) else value2
        displayed = str(cell.Text)
        error = displayed if displayed.startswith("#") else None
        normalized = normalize_value(
            error or value,
            date_system=environment.date_system,
            timezone=environment.timezone,
            formula=str(cell.Formula) if bool(cell.HasFormula) else None,
            cell_type=str(cell.NumberFormat),
            error_type=error,
        )
        return normalized
    if request.kind is ObservationKind.SHEETS:
        return normalize_value(
            [
                {"name": str(sheet.Name), "index": index, "visible": int(sheet.Visible)}
                for index, sheet in enumerate(workbook.Worksheets, start=1)
            ]
        )
    if request.kind is ObservationKind.NAMED_RANGES:
        return normalize_value(
            [{"name": str(name.Name), "refers_to": str(name.RefersTo)} for name in workbook.Names]
        )
    if request.kind is ObservationKind.EMBEDDED_SCRIPTS:
        try:
            components = workbook.VBProject.VBComponents
            return normalize_value(
                [
                    {"name": str(component.Name), "type": int(component.Type)}
                    for component in components
                ]
            )
        except Exception as exc:
            raise PermissionError("programmatic VBA project access is unavailable") from exc
    if request.kind is ObservationKind.CONTROLS_EVENTS:
        controls = []
        for sheet in workbook.Worksheets:
            for shape in sheet.Shapes:
                controls.append(
                    {
                        "sheet": str(sheet.Name),
                        "name": str(shape.Name),
                        "type": int(shape.Type),
                        "on_action": str(getattr(shape, "OnAction", "")),
                    }
                )
        return normalize_value(controls)
    if request.kind is ObservationKind.PACKAGE_HASH:
        return normalize_value(_hash_path(Path(str(workbook.FullName))))
    if request.kind is ObservationKind.ARTIFACT_INVENTORY:
        path = Path(str(workbook.FullName))
        members = sorted(zipfile.ZipFile(path).namelist()) if zipfile.is_zipfile(path) else []
        return normalize_value({"package_members": members})
    raise ValueError(f"unsupported observation: {request.kind}")


def _error(status: GateExecutionStatus, kind: str, message: str) -> OracleRunResult:
    return OracleRunResult(status=status, error={"type": kind, "message": message})


def _hash_path(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def main() -> int:
    try:
        request = json.loads(sys.stdin.read())
        response = execute_request(request)
    except Exception as exc:
        response = _error(GateExecutionStatus.FAILED, "worker_failure", str(exc))
    sys.stdout.write(response.model_dump_json() + "\n")
    return 0 if response.status is GateExecutionStatus.PASSED else 1


if __name__ == "__main__":
    raise SystemExit(main())
