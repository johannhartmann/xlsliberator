"""Docker-only operations for generated application migration candidates."""

from __future__ import annotations

import re
import tempfile
from collections.abc import Mapping
from pathlib import Path
from typing import Any, Final
from zipfile import ZIP_STORED, ZipFile

from xlsliberator.docker_runtime import (
    LIBREOFFICE_VERSION,
    DockerRuntimeUnavailable,
    LibreOfficeDockerRuntime,
)

GUI_IMAGE: Final = "xlsliberator-libreoffice-gui:26.2.4.2"
_SAFE_ID = re.compile(r"^[A-Za-z0-9_.-]{1,100}$")


def build_candidate(
    source: Path,
    candidate_bundle: Path,
    output: Path,
    *,
    timeout_seconds: int = 120,
) -> dict[str, Any]:
    """Build an ODS through a content-bound generated candidate in Docker."""
    runtime = LibreOfficeDockerRuntime(timeout_seconds=timeout_seconds)
    response = runtime.request(
        {
            "op": "build_application_target",
            "input_path": str(source),
            "candidate_path": str(candidate_bundle),
            "output_path": str(output),
            "timeout_seconds": timeout_seconds,
        }
    )
    return _require_success(response, "migration candidate build")


def run_application_scenario(
    target: Path,
    candidate_bundle: Path,
    evidence_archive: Path,
    actions: list[dict[str, Any]],
    *,
    scenario_id: str,
    adapter_config: Mapping[str, Any] | None = None,
    timeout_seconds: int = 180,
) -> dict[str, Any]:
    """Operate any generated candidate through real X11 events in Docker."""
    _safe_id(scenario_id, "scenario_id")
    runtime = LibreOfficeDockerRuntime(
        image=GUI_IMAGE,
        timeout_seconds=timeout_seconds,
    )
    response = runtime.request(
        {
            "op": "run_application_scenario",
            "ods_path": str(target),
            "candidate_path": str(candidate_bundle),
            "output_path": str(evidence_archive),
            "scenario_id": scenario_id,
            "actions": actions,
            "adapter_config": dict(adapter_config or {}),
            "timeout_seconds": timeout_seconds,
        }
    )
    return _require_success(response, "migration candidate GUI scenario")


def bundle_application_replays(
    evidence_archives: Mapping[str, Path],
    replay_archive: Path,
    *,
    replay_id: str,
    timeout_seconds: int = 180,
) -> dict[str, Any]:
    """Create one replay bundle from an arbitrary declared scenario set."""
    _safe_id(replay_id, "replay_id")
    if not 1 <= len(evidence_archives) <= 100:
        raise ValueError("replay input must contain between 1 and 100 scenarios")
    scenario_ids = list(evidence_archives)
    if len(set(scenario_ids)) != len(scenario_ids):
        raise ValueError("replay scenario IDs must be unique")
    for scenario_id in scenario_ids:
        _safe_id(scenario_id, "scenario_id")

    replay_archive.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(
        prefix="xlsliberator-application-replays-",
        suffix=".zip",
        dir=replay_archive.parent,
        delete=False,
    ) as handle:
        staged_bundle = Path(handle.name)
    try:
        with ZipFile(staged_bundle, "w", compression=ZIP_STORED) as bundle:
            for scenario_id, evidence in evidence_archives.items():
                if not evidence.is_file():
                    raise FileNotFoundError(evidence)
                bundle.write(evidence, f"{scenario_id}.zip")
        runtime = LibreOfficeDockerRuntime(
            image=GUI_IMAGE,
            timeout_seconds=timeout_seconds,
        )
        response = runtime.request(
            {
                "op": "bundle_application_replays",
                "input_path": str(staged_bundle),
                "output_path": str(replay_archive),
                "scenario_ids": scenario_ids,
                "replay_id": replay_id,
                "timeout_seconds": timeout_seconds,
            }
        )
        return _require_success(response, "migration candidate replay bundle")
    finally:
        staged_bundle.unlink(missing_ok=True)


def _safe_id(value: str, label: str) -> str:
    if _SAFE_ID.fullmatch(value) is None:
        raise ValueError(f"{label} is malformed")
    return value


def _require_success(response: dict[str, Any], operation: str) -> dict[str, Any]:
    if response.get("success"):
        data = dict(response.get("data") or {})
        if data.get("target_build") not in {None, LIBREOFFICE_VERSION}:
            raise DockerRuntimeUnavailable(f"{operation} returned the wrong target build")
        return data
    error = response.get("error") or {}
    message = str(error.get("message") or f"{operation} failed")
    traceback_text = str(error.get("traceback") or "").strip()
    container_stderr = str((response.get("data") or {}).get("container_stderr") or "").strip()
    details = [
        text
        for text in (
            traceback_text[-4_000:],
            f"container stderr:\n{container_stderr[-4_000:]}" if container_stderr else "",
        )
        if text
    ]
    detail_text = "\n".join(details)
    detail = f"\n{detail_text}" if detail_text else ""
    raise DockerRuntimeUnavailable(f"{message}{detail}")


__all__ = [
    "GUI_IMAGE",
    "build_candidate",
    "bundle_application_replays",
    "run_application_scenario",
]
