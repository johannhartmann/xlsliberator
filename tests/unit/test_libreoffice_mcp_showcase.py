from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest

from xlsliberator.libreoffice_mcp import (
    build_application_candidate,
    bundle_application_replays,
    run_application_scenario,
)


@pytest.mark.asyncio
async def test_build_showcase_tool_confines_paths(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    source = tmp_path / "source.xlsb"
    candidate = tmp_path / "candidate.zip"
    output = tmp_path / "target.ods"
    source.write_bytes(b"source")
    candidate.write_bytes(b"candidate")
    observed: dict[str, Path] = {}

    def fake_build(
        actual_source: Path,
        actual_candidate: Path,
        actual_output: Path,
    ) -> dict[str, Any]:
        observed["source"] = actual_source
        observed["candidate"] = actual_candidate
        observed["output"] = actual_output
        actual_output.write_bytes(b"ods")
        return {"status": "passed", "target_build": "26.2.4.2"}

    monkeypatch.setenv("XLSLIBERATOR_WORKSPACE_ROOTS", str(tmp_path))
    monkeypatch.setattr("xlsliberator.libreoffice_mcp.build_candidate", fake_build)

    result = await build_application_candidate(str(source), str(candidate), str(output))

    assert result["success"] is True
    assert result["operation_status"] == "passed"
    assert observed == {"source": source, "candidate": candidate, "output": output}
    assert output.read_bytes() == b"ods"


@pytest.mark.asyncio
async def test_scenario_tool_accepts_any_safe_declared_scenario(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    target = tmp_path / "target.ods"
    candidate = tmp_path / "candidate.zip"
    target.write_bytes(b"ods")
    candidate.write_bytes(b"candidate")
    monkeypatch.setenv("XLSLIBERATOR_WORKSPACE_ROOTS", str(tmp_path))
    observed: dict[str, Any] = {}

    def fake_run(*args: Any, **kwargs: Any) -> dict[str, Any]:
        observed["args"] = args
        observed["kwargs"] = kwargs
        return {"status": "passed", "target_build": "26.2.4.2"}

    monkeypatch.setattr("xlsliberator.libreoffice_mcp.run_scenario", fake_run)

    result = await run_application_scenario(
        str(target),
        str(candidate),
        str(tmp_path / "evidence.zip"),
        "invented-scenario",
        [{"kind": "observe"}],
        {"feature": "opaque"},
    )

    assert result["success"] is True
    assert result["operation_status"] == "passed"
    assert observed["args"][:2] == (target, candidate)
    assert observed["kwargs"]["scenario_id"] == "invented-scenario"
    assert observed["kwargs"]["adapter_config"] == {"feature": "opaque"}


@pytest.mark.asyncio
async def test_replay_tool_forwards_declared_set(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("XLSLIBERATOR_WORKSPACE_ROOTS", str(tmp_path))
    scenario = tmp_path / "boundary.zip"
    scenario.write_bytes(b"evidence")
    observed: dict[str, Any] = {}

    def fake_bundle(*args: Any, **kwargs: Any) -> dict[str, Any]:
        observed["args"] = args
        observed["kwargs"] = kwargs
        return {"status": "passed", "target_build": "26.2.4.2"}

    monkeypatch.setattr("xlsliberator.libreoffice_mcp.bundle_replays", fake_bundle)

    result = await bundle_application_replays(
        {"boundary": str(scenario)},
        str(tmp_path / "replay.zip"),
        "migration-42",
    )

    assert result["success"] is True
    assert result["operation_status"] == "passed"
    assert observed["args"][0] == {"boundary": scenario}
    assert observed["kwargs"]["replay_id"] == "migration-42"
