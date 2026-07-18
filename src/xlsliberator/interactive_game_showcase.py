"""Docker-only public operations for the interactive-game showcase."""

from __future__ import annotations

from pathlib import Path
from typing import Any, Final

from xlsliberator.docker_runtime import (
    DockerRuntimeUnavailable,
    LibreOfficeDockerRuntime,
)
from xlsliberator.interactive_game_uno import SOURCE_SHA256, TARGET_BUILD

GUI_IMAGE: Final = "xlsliberator-libreoffice-gui:26.2.4.2"


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
    "SOURCE_SHA256",
    "TARGET_BUILD",
    "build_target",
    "run_gui_scenario",
]
