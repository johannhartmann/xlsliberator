"""Docker-only public operations for the interactive-game showcase."""

from __future__ import annotations

import tempfile
from pathlib import Path
from typing import Any, Final, Mapping
from zipfile import ZIP_STORED, ZipFile

from xlsliberator.docker_runtime import (
    DockerRuntimeUnavailable,
    LibreOfficeDockerRuntime,
)
from xlsliberator.interactive_game_uno import SOURCE_SHA256, TARGET_BUILD

GUI_IMAGE: Final = "xlsliberator-libreoffice-gui:26.2.4.2"
PUBLIC_SCENARIOS: Final = (
    "keyboard-control",
    "timer-tick",
    "native-controls",
    "document-events",
    "line-collapse",
)


def build_target(
    source: Path,
    output: Path,
    *,
    timeout_seconds: int = 120,
) -> dict[str, Any]:
    """Build the target in the pinned stock LibreOffice Docker image."""
    runtime = LibreOfficeDockerRuntime(timeout_seconds=timeout_seconds)
    response = runtime.request(
        {
            "op": "build_interactive_game_target",
            "input_path": str(source),
            "output_path": str(output),
            "timeout_seconds": timeout_seconds,
        }
    )
    return _require_success(response, "interactive-game target build")


def run_gui_scenario(
    target: Path,
    evidence_archive: Path,
    actions: list[dict[str, Any]],
    *,
    scenario_id: str | None = None,
    timer_enabled: bool = True,
    timeout_seconds: int = 180,
) -> dict[str, Any]:
    """Operate the target through real X11 events in the GUI Docker image."""
    runtime = LibreOfficeDockerRuntime(
        image=GUI_IMAGE,
        timeout_seconds=timeout_seconds,
    )
    response = runtime.request(
        {
            "op": "run_gui_scenario",
            "ods_path": str(target),
            "output_path": str(evidence_archive),
            "scenario_id": scenario_id or evidence_archive.stem,
            "adapter": "interactive-game",
            "actions": actions,
            "timer_enabled": timer_enabled,
            "timeout_seconds": timeout_seconds,
        }
    )
    return _require_success(response, "interactive-game GUI scenario")


def bundle_gui_replays(
    evidence_archives: Mapping[str, Path],
    replay_archive: Path,
    *,
    timeout_seconds: int = 180,
) -> dict[str, Any]:
    """Create one public replay from all canonical GUI scenario evidence."""
    if set(evidence_archives) != set(PUBLIC_SCENARIOS):
        raise ValueError("replay input must contain every canonical public scenario exactly once")
    replay_archive.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(
        prefix="xlsliberator-showcase-replays-",
        suffix=".zip",
        dir=replay_archive.parent,
        delete=False,
    ) as handle:
        staged_bundle = Path(handle.name)
    try:
        with ZipFile(staged_bundle, "w", compression=ZIP_STORED) as bundle:
            for scenario_id in PUBLIC_SCENARIOS:
                evidence = evidence_archives[scenario_id]
                if not evidence.is_file():
                    raise FileNotFoundError(evidence)
                bundle.write(evidence, f"{scenario_id}.zip")
        runtime = LibreOfficeDockerRuntime(
            image=GUI_IMAGE,
            timeout_seconds=timeout_seconds,
        )
        response = runtime.request(
            {
                "op": "bundle_gui_replays",
                "input_path": str(staged_bundle),
                "output_path": str(replay_archive),
                "timeout_seconds": timeout_seconds,
            }
        )
        return _require_success(response, "interactive-game replay bundle")
    finally:
        staged_bundle.unlink(missing_ok=True)


def _require_success(response: dict[str, Any], operation: str) -> dict[str, Any]:
    if response.get("success"):
        data = dict(response.get("data") or {})
        if data.get("target_build") not in {None, TARGET_BUILD}:
            raise DockerRuntimeUnavailable(f"{operation} returned the wrong target build")
        return data
    error = response.get("error") or {}
    message = str(error.get("message") or f"{operation} failed")
    traceback_text = str(error.get("traceback") or "").strip()
    detail = f"\n{traceback_text[-4_000:]}" if traceback_text else ""
    raise DockerRuntimeUnavailable(f"{message}{detail}")


__all__ = [
    "GUI_IMAGE",
    "PUBLIC_SCENARIOS",
    "SOURCE_SHA256",
    "TARGET_BUILD",
    "build_target",
    "bundle_gui_replays",
    "run_gui_scenario",
]
