"""Fail-closed VBA to Python-UNO translation facade."""

from __future__ import annotations

import os
import re
from dataclasses import dataclass, field
from pathlib import Path

from xlsliberator.docker_runtime import DockerRuntimeUnavailable, LibreOfficeDockerRuntime
from xlsliberator.translation_service import (
    AnthropicTranslationProvider,
    ProjectTranslationResult,
    TranslationProvider,
    TranslationProviderUnavailable,
    TranslationService,
    TranslationStatus,
)


class VBATranslationError(Exception):
    """Raised by compatibility callers that require an accepted translation."""


@dataclass
class TranslationResult:
    """Backward-compatible result enriched with canonical translation evidence."""

    python_code: str
    warnings: list[str]
    unsupported_features: list[str]
    status: TranslationStatus = TranslationStatus.FAILED
    provenance: dict[str, object] = field(default_factory=dict)
    evidence_manifest: str | None = None

    @property
    def succeeded(self) -> bool:
        """Return true only for an accepted, fully proven translation."""
        return self.status is TranslationStatus.SUCCEEDED and bool(self.python_code)


def translate_vba_to_python(
    vba_code: str,
    module_name: str | None = None,
    *,
    provider: TranslationProvider | None = None,
    cache_path: Path | None = None,
    evidence_dir: Path | None = None,
) -> TranslationResult:
    """Translate one VBA module through the typed project service.

    LibreOffice/PyUNO provenance is resolved through the pinned Docker image.
    No host office executable or host UNO module is inspected or imported.
    """
    name = module_name or "Module1"
    project = translate_vba_project(
        {name: vba_code},
        provider=provider,
        cache_path=cache_path,
        evidence_dir=evidence_dir,
    )
    result = _compatibility_result(project, name)
    if result.python_code:
        result.python_code = _inject_source_markers(
            result.python_code,
            _extract_vba_procedure_names(vba_code),
            module_name,
        )
    return result


def translate_vba_project(
    modules: dict[str, str],
    *,
    provider: TranslationProvider | None = None,
    cache_path: Path | None = None,
    evidence_dir: Path | None = None,
    event_targets: tuple[str, ...] = (),
) -> ProjectTranslationResult:
    """Translate a complete project using one context and one runtime identity probe."""
    try:
        selected_provider = provider or _configured_provider()
    except TranslationProviderUnavailable as exc:
        return ProjectTranslationResult(
            status=TranslationStatus.UNAVAILABLE,
            runtime_status=TranslationStatus.UNAVAILABLE,
            errors=[str(exc)],
        )

    runtime_id: str | None = None
    package_manifest: list[dict[str, object]] = []
    runtime_error: str | None
    try:
        identity = LibreOfficeDockerRuntime().resolve_identity()
        runtime_id = identity.image_id
        package_manifest = [
            dict(item)
            for item in identity.probe.get("installed_package_manifest", [])
            if isinstance(item, dict)
        ]
    except DockerRuntimeUnavailable as exc:
        runtime_error = f"target runtime provenance unavailable: {exc}"
    else:
        runtime_error = None

    service = TranslationService(
        selected_provider,
        cache_path=cache_path,
        mapping_paths=(Path("rules/vba_api_map.yaml"), Path("rules/event_map.yaml")),
        target_runtime_image_id=runtime_id,
        target_package_manifest=package_manifest,
        evidence_dir=evidence_dir,
    )
    project = service.translate_project(modules, event_targets=event_targets)
    if runtime_error:
        project.errors.append(runtime_error)
    if project.accepted:
        for name, module in project.modules.items():
            if module.python_code:
                module.python_code = _inject_source_markers(
                    module.python_code,
                    _extract_vba_procedure_names(modules[name]),
                    name,
                )
    return project


def _configured_provider() -> TranslationProvider:
    provider_name = os.environ.get("XLSLIBERATOR_TRANSLATION_PROVIDER", "anthropic").lower()
    model_id = os.environ.get("XLSLIBERATOR_TRANSLATION_MODEL", "claude-sonnet-4-5")
    if provider_name != "anthropic":
        raise TranslationProviderUnavailable(
            f"translation provider {provider_name!r} is not configured in this process"
        )
    return AnthropicTranslationProvider(model_id=model_id)


def _compatibility_result(project: ProjectTranslationResult, module_name: str) -> TranslationResult:
    module = project.modules.get(module_name)
    code = module.python_code if module and project.accepted else ""
    provenance = module.provenance.model_dump(mode="json") if module else {}
    warnings = [*project.warnings, *project.errors]
    if module and module.error:
        warnings.append(module.error)
    return TranslationResult(
        python_code=code or "",
        warnings=warnings,
        unsupported_features=[],
        status=project.status,
        provenance=provenance,
        evidence_manifest=project.evidence_manifest,
    )


def _extract_vba_procedure_names(vba_code: str) -> list[str]:
    return [
        match.group(1)
        for match in re.finditer(
            r"^[ \t]*(?:(?:Public|Private|Friend|Static)[ \t]+)*(?:Sub|Function)[ \t]+(\w+)",
            vba_code,
            re.IGNORECASE | re.MULTILINE,
        )
    ]


def _source_marker(module_name: str | None, procedure: str) -> str:
    module = module_name or "unknown"
    artifact_id = f"{module}.{procedure}"
    return (
        f"# xlsliberator-source: module={module}; procedure={procedure}; artifact_id={artifact_id}"
    )


def _inject_source_markers(
    python_code: str,
    procedures: list[str],
    module_name: str | None,
) -> str:
    """Inject source-map markers after matching generated function definitions."""
    updated = python_code
    for procedure in procedures:
        pattern = rf"(def\s+{re.escape(procedure)}\s*\(.*\)\s*(?:->[^\n:]+)?:[ \t]*\n)"
        replacement = rf"\1    {_source_marker(module_name, procedure)}\n"
        updated = re.sub(pattern, replacement, updated, count=1)
    return updated


def create_event_handler_stub(
    event_name: str,
    vba_code: str,
    *,
    provider: TranslationProvider | None = None,
) -> str:
    """Return an accepted event translation or raise a typed compatibility error."""
    result = translate_vba_to_python(vba_code, event_name, provider=provider)
    if not result.succeeded:
        detail = "; ".join(result.warnings) or result.status.value
        raise VBATranslationError(f"Event handler translation was not accepted: {detail}")
    return result.python_code
