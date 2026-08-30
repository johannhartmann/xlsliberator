from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace
from typing import Any

import pytest

from xlsliberator.libreoffice_mcp import (
    build_application_candidate,
    bundle_application_replays,
    run_application_scenario,
    run_declared_acceptance,
    run_declared_mutation_suite,
)
from xlsliberator.validation_models import GateExecutionStatus


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


@pytest.mark.asyncio
async def test_declared_acceptance_runs_at_trusted_boundary(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    target = tmp_path / "target.ods"
    acceptance = tmp_path / "acceptance.yaml"
    target.write_bytes(b"ods")
    acceptance.write_text("acceptance", encoding="utf-8")
    observed: dict[str, Any] = {}

    def fake_load(path: Path) -> object:
        observed["acceptance_path"] = path
        return object()

    def fake_run(**kwargs: Any) -> Any:
        observed.update(kwargs)
        return SimpleNamespace(
            status=GateExecutionStatus.PASSED,
            model_dump=lambda **_kwargs: {"evidence_id": "evidence-1"},
        )

    monkeypatch.setenv("XLSLIBERATOR_WORKSPACE_ROOTS", str(tmp_path))
    monkeypatch.setattr("xlsliberator.libreoffice_mcp.load_acceptance", fake_load)
    monkeypatch.setattr("xlsliberator.libreoffice_mcp.run_acceptance", fake_run)
    monkeypatch.setattr(
        "xlsliberator.libreoffice_mcp.LibreOfficeSessionScenarioRunner",
        lambda **_kwargs: object(),
    )

    result = await run_declared_acceptance(
        str(target),
        str(acceptance),
        str(tmp_path / "public-evidence"),
    )

    assert result["operation_status"] == "passed"
    assert result["evidence_id"] == "evidence-1"
    assert observed["acceptance_path"] == acceptance
    assert observed["workbook"] == target
    assert observed["evidence_dir"] == tmp_path / "public-evidence"


@pytest.mark.asyncio
async def test_declared_mutation_suite_preserves_failed_status(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    target = tmp_path / "target.ods"
    acceptance = tmp_path / "acceptance.yaml"
    target.write_bytes(b"ods")
    acceptance.write_text("acceptance", encoding="utf-8")

    monkeypatch.setenv("XLSLIBERATOR_WORKSPACE_ROOTS", str(tmp_path))
    monkeypatch.setattr("xlsliberator.libreoffice_mcp.load_acceptance", lambda _path: object())
    monkeypatch.setattr(
        "xlsliberator.libreoffice_mcp.LibreOfficeSessionScenarioRunner",
        lambda **_kwargs: object(),
    )
    monkeypatch.setattr(
        "xlsliberator.libreoffice_mcp.run_mutation_campaign",
        lambda **_kwargs: SimpleNamespace(
            status=GateExecutionStatus.FAILED,
            model_dump=lambda **_kwargs: {"cases": [{"outcome": "survived"}]},
        ),
    )

    result = await run_declared_mutation_suite(
        str(target),
        str(acceptance),
        str(tmp_path / "mutations"),
    )

    assert result["success"] is False
    assert result["operation_status"] == "failed"
    assert result["cases"] == [{"outcome": "survived"}]
