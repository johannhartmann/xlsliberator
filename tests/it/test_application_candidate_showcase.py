"""Real pinned-LibreOffice acceptance for a generated application candidate."""

from __future__ import annotations

import hashlib
import json
import os
from pathlib import Path
from zipfile import ZipFile

import pytest

from xlsliberator.application_showcase import (
    build_candidate,
    bundle_application_replays,
    run_application_scenario,
)
from xlsliberator.candidate_runtime import package_candidate_directory

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


def test_complete_interactive_game_acceptance_in_real_gui_runtime(tmp_path: Path) -> None:
    source = Path("demos/interactive-game/source/TetrisGameDemo.xlsb")
    candidate_directory = Path("demos/interactive-game/candidate")
    evidence_root = _evidence_root(tmp_path)
    candidate_bundle = evidence_root / "migration-candidate.zip"
    target = evidence_root / "interactive-game.ods"
    source_before = _sha256(source)
    source_manifest = json.loads(
        (candidate_directory / "manifest.json").read_text(encoding="utf-8")
    )

    package = package_candidate_directory(candidate_directory, candidate_bundle)
    build = build_candidate(source, candidate_bundle, target)

    assert source_before == source_manifest["source_sha256"] == _sha256(source)
    assert package["candidate_id"] == "interactive-game-source-derived"
    assert build["target_build"] == "26.2.4.2"
    assert build["source_sha256"] == source_before
    assert build["candidate_bundle_sha256"] == _sha256(candidate_bundle)
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
        result = run_application_scenario(
            target,
            candidate_bundle,
            evidence,
            list(scenario["actions"]),  # type: ignore[arg-type]
            scenario_id=scenario_id,
            adapter_config=dict(scenario["adapter_config"]),  # type: ignore[arg-type]
        )
        results[scenario_id] = result
        assert result["status"] == "passed"
        assert result["event_layer"] == "xvfb-openbox-xdotool"
        assert result["target_sha256"] == target_sha256
        assert result["candidate_id"] == package["candidate_id"]
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
    replay = bundle_application_replays(
        evidence_archives,
        replay_archive,
        replay_id="interactive-game",
    )
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
