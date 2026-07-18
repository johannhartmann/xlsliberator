"""Deterministic, provider-neutral workbook migration primitives."""

from __future__ import annotations

import zipfile
from collections.abc import Mapping
from pathlib import Path
from typing import Any, Literal

from pydantic import BaseModel, ConfigDict, Field

from xlsliberator.container_boundary import require_application_container
from xlsliberator.docker_runtime import DockerRuntimeUnavailable, LibreOfficeDockerRuntime
from xlsliberator.embed_macros import embed_python_macros
from xlsliberator.extract_vba import VBAModuleIR, extract_vba_modules
from xlsliberator.inspect_workbook import inspect_workbook
from xlsliberator.libreoffice_scenario_runner import LibreOfficeScenarioRunner
from xlsliberator.scenarios.models import (
    Action,
    ActionKind,
    EnvironmentManifest,
    RuntimeTrace,
    Scenario,
    ScenarioStep,
)
from xlsliberator.validation_models import GateExecutionStatus, WorkbookArtifactIR


class PrimitiveResult(BaseModel):
    """Common truthful result fields for deterministic primitives."""

    model_config = ConfigDict(extra="forbid", arbitrary_types_allowed=True)

    status: GateExecutionStatus
    errors: list[str] = Field(default_factory=list)
    warnings: list[str] = Field(default_factory=list)
    evidence: dict[str, Any] = Field(default_factory=dict)

    @property
    def success(self) -> bool:
        """Return true only when the operation passed."""
        return self.status is GateExecutionStatus.PASSED


class WorkbookInspectionResult(PrimitiveResult):
    """Typed source or target workbook inspection result."""

    inventory: WorkbookArtifactIR | None = None


class NativeConversionResult(PrimitiveResult):
    """Typed result from the pinned LibreOffice native converter."""

    output_path: str | None = None
    runtime_identity: dict[str, Any] = Field(default_factory=dict)


class VBAProjectExtractionResult(PrimitiveResult):
    """Typed raw VBA project extraction result."""

    modules: list[VBAModuleIR] = Field(default_factory=list)


class ScriptUpsertResult(PrimitiveResult):
    """Typed result from deterministic Python-script package upsert."""

    output_path: str | None = None
    module_names: list[str] = Field(default_factory=list)
    unresolved_bindings: list[str] = Field(default_factory=list)


class PackageValidationResult(PrimitiveResult):
    """Typed structural ODS package validation result."""

    package_path: str
    member_count: int = 0
    mimetype_first: bool = False
    mimetype_stored: bool = False


class AcceptanceScenarioResult(PrimitiveResult):
    """Typed target-runtime scenario result."""

    trace: RuntimeTrace | None = None


def inspect_source_workbook(path: str | Path) -> WorkbookInspectionResult:
    """Inspect a source workbook without model calls or office execution."""
    return _inspect(path, role="source")


def inspect_target_ods(path: str | Path) -> WorkbookInspectionResult:
    """Inspect a target ODS without model calls or office execution."""
    return _inspect(path, role="target")


def _inspect(
    path: str | Path,
    *,
    role: Literal["source", "target"],
) -> WorkbookInspectionResult:
    try:
        inventory = inspect_workbook(Path(path), role=role)
    except Exception as exc:
        return WorkbookInspectionResult(
            status=GateExecutionStatus.FAILED,
            errors=[str(exc)],
        )
    return WorkbookInspectionResult(
        status=GateExecutionStatus.PASSED,
        inventory=inventory,
    )


def native_convert_workbook(
    input_path: str | Path,
    output_path: str | Path,
    *,
    timeout_seconds: int = 120,
) -> NativeConversionResult:
    """Convert with the pinned Docker office runtime and return its identity."""
    source = Path(input_path)
    destination = Path(output_path)
    try:
        require_application_container()
        response = LibreOfficeDockerRuntime(timeout_seconds=timeout_seconds).convert(
            source,
            destination,
        )
    except DockerRuntimeUnavailable as exc:
        return NativeConversionResult(
            status=GateExecutionStatus.UNAVAILABLE,
            errors=[str(exc)],
        )
    except Exception as exc:
        return NativeConversionResult(
            status=GateExecutionStatus.FAILED,
            errors=[str(exc)],
        )
    if not destination.is_file():
        return NativeConversionResult(
            status=GateExecutionStatus.FAILED,
            errors=["LibreOffice conversion did not produce an output file"],
        )
    data = dict(response.get("data") or {})
    return NativeConversionResult(
        status=GateExecutionStatus.PASSED,
        output_path=str(destination),
        runtime_identity={
            "image_id": data.get("container_image_id"),
            "image_reference": data.get("container_image"),
        },
        evidence={"worker_response": response},
    )


def extract_vba_project(path: str | Path) -> VBAProjectExtractionResult:
    """Extract complete raw VBA modules and boundaries without interpreting them."""
    try:
        modules = extract_vba_modules(path)
    except Exception as exc:
        return VBAProjectExtractionResult(
            status=GateExecutionStatus.FAILED,
            errors=[str(exc)],
        )
    return VBAProjectExtractionResult(
        status=GateExecutionStatus.PASSED,
        modules=modules,
    )


def upsert_python_modules(
    ods_path: str | Path,
    modules: Mapping[str, str],
) -> ScriptUpsertResult:
    """Upsert agent-produced target-native modules into an existing ODS."""
    path = Path(ods_path)
    before = validate_ods_package(path)
    if not before.success:
        return ScriptUpsertResult(
            status=GateExecutionStatus.FAILED,
            errors=before.errors,
        )
    if not modules:
        return ScriptUpsertResult(
            status=GateExecutionStatus.NOT_RUN,
            errors=["No Python modules were supplied"],
        )
    try:
        unresolved = embed_python_macros(path, dict(modules))
    except Exception as exc:
        return ScriptUpsertResult(
            status=GateExecutionStatus.FAILED,
            errors=[str(exc)],
        )
    after = validate_ods_package(path)
    if not after.success:
        return ScriptUpsertResult(
            status=GateExecutionStatus.FAILED,
            errors=after.errors,
        )
    unresolved_names = [binding.source_handler for binding in unresolved]
    return ScriptUpsertResult(
        status=(GateExecutionStatus.FAILED if unresolved_names else GateExecutionStatus.PASSED),
        output_path=str(path),
        module_names=sorted(modules),
        unresolved_bindings=unresolved_names,
        errors=(
            [f"Unresolved event binding: {name}" for name in unresolved_names]
            if unresolved_names
            else []
        ),
    )


def validate_ods_package(path: str | Path) -> PackageValidationResult:
    """Validate the structural invariants required of an ODS ZIP package."""
    package = Path(path)
    errors: list[str] = []
    member_count = 0
    mimetype_first = False
    mimetype_stored = False
    if not package.is_file():
        errors.append("ODS package does not exist")
    elif not zipfile.is_zipfile(package):
        errors.append("ODS package is not a ZIP archive")
    else:
        try:
            with zipfile.ZipFile(package) as archive:
                infos = archive.infolist()
                member_count = len(infos)
                names = [info.filename for info in infos]
                if len(names) != len(set(names)):
                    errors.append("ODS package contains duplicate member paths")
                if infos:
                    mimetype_first = infos[0].filename == "mimetype"
                    mimetype_stored = (
                        mimetype_first and infos[0].compress_type == zipfile.ZIP_STORED
                    )
                if not mimetype_first:
                    errors.append("ODS mimetype entry is missing or not first")
                elif not mimetype_stored:
                    errors.append("ODS mimetype entry must be uncompressed")
                else:
                    value = archive.read("mimetype")
                    expected = b"application/vnd.oasis.opendocument.spreadsheet"
                    if value != expected:
                        errors.append("ODS mimetype value is invalid")
                if "META-INF/manifest.xml" not in names:
                    errors.append("ODS manifest is missing")
                bad_paths = [
                    name
                    for name in names
                    if name.startswith("/") or ".." in Path(name).parts or "\\" in name
                ]
                if bad_paths:
                    errors.append("ODS package contains unsafe member paths")
                corrupt = archive.testzip()
                if corrupt is not None:
                    errors.append(f"ODS package member failed CRC validation: {corrupt}")
        except (OSError, zipfile.BadZipFile, KeyError) as exc:
            errors.append(str(exc))
    return PackageValidationResult(
        status=GateExecutionStatus.FAILED if errors else GateExecutionStatus.PASSED,
        package_path=str(package),
        member_count=member_count,
        mimetype_first=mimetype_first,
        mimetype_stored=mimetype_stored,
        errors=errors,
    )


def run_acceptance_scenario(
    ods_path: str | Path,
    scenario: Scenario,
    environment: EnvironmentManifest | None = None,
) -> AcceptanceScenarioResult:
    """Run one declared scenario in the pinned target runtime."""
    trace = LibreOfficeScenarioRunner().run(
        Path(ods_path),
        environment or EnvironmentManifest(),
        scenario,
    )
    return AcceptanceScenarioResult(
        status=trace.status,
        trace=trace,
        errors=(
            [str((trace.error or {}).get("message") or "scenario failed")]
            if trace.status is not GateExecutionStatus.PASSED
            else []
        ),
    )


def run_target_lifecycle(
    ods_path: str | Path,
    *,
    macro_uri: str | None = None,
    environment: EnvironmentManifest | None = None,
) -> AcceptanceScenarioResult:
    """Open, recalculate, optionally run a macro, save, close, and reopen."""
    steps = [
        ScenarioStep(id="open", action=Action(kind=ActionKind.OPEN)),
        ScenarioStep(id="recalculate", action=Action(kind=ActionKind.RECALCULATE)),
    ]
    if macro_uri is not None:
        steps.append(
            ScenarioStep(
                id="execute-python-macro",
                action=Action(
                    kind=ActionKind.INVOKE_MACRO,
                    parameters={"script_uri": macro_uri},
                ),
            )
        )
    steps.extend(
        [
            ScenarioStep(id="save", action=Action(kind=ActionKind.SAVE)),
            ScenarioStep(id="close", action=Action(kind=ActionKind.CLOSE)),
            ScenarioStep(id="reopen", action=Action(kind=ActionKind.REOPEN)),
        ]
    )
    return run_acceptance_scenario(
        ods_path,
        Scenario(
            id="target-lifecycle",
            description="Deterministic target lifecycle smoke",
            steps=steps,
        ),
        environment,
    )
