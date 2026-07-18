from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest

from xlsliberator.libreoffice_mcp import (
    build_interactive_game_target,
    bundle_interactive_game_replays,
    run_interactive_game_scenario,
)


@pytest.mark.asyncio
async def test_build_showcase_tool_confines_paths(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    source = tmp_path / "source.xlsb"
    output = tmp_path / "target.ods"
    source.write_bytes(b"source")
    observed: dict[str, Path] = {}

    def fake_build(actual_source: Path, actual_output: Path) -> dict[str, Any]:
        observed["source"] = actual_source
        observed["output"] = actual_output
        actual_output.write_bytes(b"ods")
        return {"status": "passed", "target_build": "26.2.4.2"}

    monkeypatch.setenv("XLSLIBERATOR_WORKSPACE_ROOTS", str(tmp_path))
    monkeypatch.setattr("xlsliberator.libreoffice_mcp.build_target", fake_build)

    result = await build_interactive_game_target(str(source), str(output))

    assert result["success"] is True
    assert result["operation_status"] == "passed"
    assert observed == {"source": source, "output": output}
    assert output.read_bytes() == b"ods"


@pytest.mark.asyncio
async def test_scenario_tool_rejects_noncanonical_scenario(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    target = tmp_path / "target.ods"
    target.write_bytes(b"ods")
    monkeypatch.setenv("XLSLIBERATOR_WORKSPACE_ROOTS", str(tmp_path))

    result = await run_interactive_game_scenario(
        str(target),
        str(tmp_path / "evidence.zip"),
        "invented-scenario",
        [],
    )

    assert result["success"] is False
    assert result["operation_status"] == "failed"
    assert result["error"]["type"] == "ValueError"


@pytest.mark.asyncio
async def test_replay_tool_requires_exact_canonical_set(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("XLSLIBERATOR_WORKSPACE_ROOTS", str(tmp_path))

    result = await bundle_interactive_game_replays(
        {"keyboard-control": str(tmp_path / "keyboard.zip")},
        str(tmp_path / "replay.zip"),
    )

    assert result["success"] is False
    assert result["operation_status"] == "failed"
    assert result["error"]["type"] == "ValueError"
