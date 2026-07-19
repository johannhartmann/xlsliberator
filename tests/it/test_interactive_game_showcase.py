"""Real pinned-LibreOffice and X11 acceptance for the interactive game."""

from __future__ import annotations

import hashlib
import json
import os
from pathlib import Path
from zipfile import ZipFile

import pytest

from xlsliberator.docker_runtime import LibreOfficeDockerRuntime
from xlsliberator.interactive_game_showcase import (
    GUI_IMAGE,
    build_target,
    bundle_gui_replays,
    run_gui_scenario,
)
from xlsliberator.interactive_game_uno import SOURCE_SHA256

pytestmark = [pytest.mark.integration, pytest.mark.docker]


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _scenario(name: str) -> dict[str, object]:
    path = Path("demos/interactive-game/showcase/scenarios") / f"{name}.json"
    return json.loads(path.read_text(encoding="utf-8"))


def _evidence_root(tmp_path: Path) -> Path:
    configured = os.environ.get("XLSLIBERATOR_SHOWCASE_ARTIFACT_DIR")
    root = Path(configured) if configured else tmp_path
    root.mkdir(parents=True, exist_ok=True)
    return root


def test_minimal_calc_document_opens_in_real_gui_runtime(tmp_path: Path) -> None:
    """Prove the pinned GUI runtime can expose and inspect a basic Calc view."""
    source = Path("tests/fixtures/scenarios/basic.ods")
    evidence = _evidence_root(tmp_path) / "minimal-calc.zip"
    runtime = LibreOfficeDockerRuntime(image=GUI_IMAGE, timeout_seconds=180)

    response = runtime.request(
        {
            "op": "run_gui_scenario",
            "ods_path": str(source),
            "output_path": str(evidence),
            "scenario_id": "minimal-calc",
            "actions": [
                {
                    "kind": "observe",
                    "sheet": "Sheet1",
                    "address": "A1",
                    "expect_value": 2,
                    "state_cells": [{"sheet": "Sheet1", "address": "A1"}],
                }
            ],
            "timer_enabled": False,
            "timeout_seconds": 180,
        }
    )

    assert response.get("success"), response
    result = response["data"]
    assert result["status"] == "passed"
    assert result["scenario_id"] == "minimal-calc"
    assert result["controller_sessions"] == []
    assert result["operations"][0]["result"]["value"] == 2
    assert evidence.is_file()
    with ZipFile(evidence) as archive:
        assert {
            "recording.webm",
            "replay.html",
            "result.json",
            "working-copy.ods",
        } <= set(archive.namelist())


def test_complete_interactive_game_acceptance_in_real_gui_runtime(tmp_path: Path) -> None:
    source = Path("demos/interactive-game/source/TetrisGameDemo.xlsb")
    evidence_root = _evidence_root(tmp_path)
    target = evidence_root / "interactive-game.ods"
    source_before = _sha256(source)

    build = build_target(source, target)

    assert source_before == SOURCE_SHA256 == _sha256(source)
    assert build["target_build"] == "26.2.4.2"
    assert build["source_sha256"] == SOURCE_SHA256
    assert build["embedded_script_bindings"] == 0
    assert build["control_lifecycle"] == "docker-runtime-native"
    assert target.is_file()
    target_sha256 = _sha256(target)
    with ZipFile(target) as archive:
        names = set(archive.namelist())
        assert not any(name.startswith(("Basic/", "Scripts/")) for name in names)
        assert not any(name.endswith("vbaProject.bin") for name in names)
        content = archive.read("content.xml")
        assert b"script:event-listener" not in content
        assert b"com.sun.star.form.component.CommandButton" not in content
        assert b"GameStart" in content
        assert b"GamePause" in content
        assert b"GameReset" in content

    results: dict[str, dict[str, object]] = {}
    evidence_archives: dict[str, Path] = {}
    operation_count = 0
    for scenario_id in (
        "keyboard-control",
        "timer-tick",
        "native-controls",
        "document-events",
        "line-collapse",
    ):
        scenario = _scenario(scenario_id)
        evidence = evidence_root / f"{scenario_id}.zip"
        evidence_archives[scenario_id] = evidence
        result = run_gui_scenario(
            target,
            evidence,
            list(scenario["actions"]),  # type: ignore[arg-type]
            timer_enabled=bool(scenario["timer_enabled"]),
        )
        results[scenario_id] = result
        assert result["status"] == "passed"
        assert result["event_layer"] == "xvfb-openbox-xdotool"
        assert result["source_sha256"] == target_sha256
        operations = result["operations"]
        assert isinstance(operations, list) and operations
        operation_count += len(operations)
        assert evidence.is_file()
        with ZipFile(evidence) as archive:
            names = set(archive.namelist())
            assert "result.json" in names
            assert "recording.webm" in names
            assert "replay.html" in names
            assert any(name.endswith(".png") for name in names)
            assert "working-copy.ods" in names

    replay_archive = evidence_root / "public-showcase-replay.zip"
    replay = bundle_gui_replays(evidence_archives, replay_archive)
    assert replay["status"] == "passed"
    assert replay["target_sha256"] == target_sha256
    assert set(replay["covered_scenarios"]) == set(evidence_archives)
    assert replay["operation_count"] == operation_count
    with ZipFile(replay_archive) as archive:
        assert set(archive.namelist()) == {
            "public/replay/events.json",
            "public/replay/index.html",
            "public/replay/showcase.webm",
        }
        events = json.loads(archive.read("public/replay/events.json"))
        assert events["status"] == "passed"
        assert events["scenario_id"] == "interactive-game"
        assert events["target_sha256"] == target_sha256
        assert len(events["operations"]) == replay["operation_count"]
        assert archive.read("public/replay/showcase.webm").startswith(b"\x1aE\xdf\xa3")

    timer_sessions = results["timer-tick"]["controller_sessions"]
    assert isinstance(timer_sessions, list) and len(timer_sessions) == 1
    timer_events = timer_sessions[0]["events"]
    third_tick = [event for event in timer_events if event["kind"] == "timer"][2]
    high_scores = next(
        event
        for event in timer_events
        if event["kind"] == "control" and event["control_name"] == "GameHighScores"
    )
    assert high_scores["sequence"] < third_tick["sequence"]

    document_sessions = results["document-events"]["controller_sessions"]
    assert isinstance(document_sessions, list) and len(document_sessions) == 2
    assert all(session["control_bindings"] == 5 for session in document_sessions)
    assert all(session["key_handler_installed"] for session in document_sessions)
    assert _sha256(source) == source_before
