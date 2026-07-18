"""Office-free contract tests for the Docker-only interactive-game boundary."""

from __future__ import annotations

import hashlib
import json
from pathlib import Path
from typing import Any

import pytest

from xlsliberator.interactive_game_engine import state_from_json
from xlsliberator.interactive_game_showcase import (
    GUI_IMAGE,
    PUBLIC_SCENARIOS,
    build_target,
    bundle_gui_replays,
    run_gui_scenario,
)
from xlsliberator.interactive_game_uno import SOURCE_SHA256


class _FakeRuntime:
    instances: list[_FakeRuntime] = []

    def __init__(self, image: str | None = None, **_kwargs: Any) -> None:
        self.image = image
        self.payload: dict[str, Any] | None = None
        self.instances.append(self)

    def request(self, payload: dict[str, Any]) -> dict[str, Any]:
        self.payload = payload
        return {
            "success": True,
            "data": {
                "status": "passed",
                "target_build": "26.2.4.2",
            },
        }


def test_build_target_uses_only_the_pinned_worker_boundary(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    source = tmp_path / "source.xlsb"
    source.write_bytes(b"source")
    output = tmp_path / "target.ods"
    monkeypatch.setattr(
        "xlsliberator.interactive_game_showcase.LibreOfficeDockerRuntime",
        _FakeRuntime,
    )
    _FakeRuntime.instances.clear()

    result = build_target(source, output)

    assert result["status"] == "passed"
    assert len(_FakeRuntime.instances) == 1
    runtime = _FakeRuntime.instances[0]
    assert runtime.image is None
    assert runtime.payload == {
        "op": "build_interactive_game_target",
        "input_path": str(source),
        "output_path": str(output),
        "timeout_seconds": 120,
    }


def test_gui_scenario_selects_dedicated_image_and_forwards_timer_policy(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    target = tmp_path / "target.ods"
    target.write_bytes(b"target")
    output = tmp_path / "evidence.zip"
    actions = [{"kind": "observe", "sheet": "game", "address": "C2"}]
    monkeypatch.setattr(
        "xlsliberator.interactive_game_showcase.LibreOfficeDockerRuntime",
        _FakeRuntime,
    )
    _FakeRuntime.instances.clear()

    run_gui_scenario(target, output, actions, timer_enabled=False)

    runtime = _FakeRuntime.instances[0]
    assert runtime.image == GUI_IMAGE
    assert runtime.payload is not None
    assert runtime.payload["op"] == "run_gui_scenario"
    assert runtime.payload["adapter"] == "interactive-game"
    assert runtime.payload["timer_enabled"] is False
    assert runtime.payload["actions"] == actions


def test_replay_bundle_selects_gui_image_and_requires_all_public_scenarios(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    evidence: dict[str, Path] = {}
    for scenario_id in PUBLIC_SCENARIOS:
        path = tmp_path / f"{scenario_id}.zip"
        path.write_bytes(scenario_id.encode())
        evidence[scenario_id] = path
    output = tmp_path / "replay.zip"
    monkeypatch.setattr(
        "xlsliberator.interactive_game_showcase.LibreOfficeDockerRuntime",
        _FakeRuntime,
    )
    _FakeRuntime.instances.clear()

    bundle_gui_replays(evidence, output)

    runtime = _FakeRuntime.instances[0]
    assert runtime.image == GUI_IMAGE
    assert runtime.payload is not None
    assert runtime.payload["op"] == "bundle_gui_replays"
    assert runtime.payload["output_path"] == str(output)
    assert not Path(runtime.payload["input_path"]).exists()

    with pytest.raises(ValueError, match="every canonical public scenario"):
        bundle_gui_replays(
            {scenario: path for scenario, path in evidence.items() if scenario != "timer-tick"},
            output,
        )


def test_public_scenarios_have_exact_ids_and_valid_game_fixtures() -> None:
    root = Path("demos/interactive-game/showcase/scenarios")
    scenarios = [
        json.loads(path.read_text(encoding="utf-8")) for path in sorted(root.glob("*.json"))
    ]

    assert {scenario["scenario_id"] for scenario in scenarios} == {
        "keyboard-control",
        "timer-tick",
        "native-controls",
        "document-events",
        "line-collapse",
    }
    for scenario in scenarios:
        assert scenario["schema_version"] == "1.0.0"
        assert scenario["actions"]
        for action in scenario["actions"]:
            if action["kind"] == "load_game_state":
                state_from_json(action["state_json"])


def test_source_identity_is_bound_to_the_real_public_workbook() -> None:
    source = Path("demos/interactive-game/source/TetrisGameDemo.xlsb")

    assert source.is_file()
    assert hashlib.sha256(source.read_bytes()).hexdigest() == SOURCE_SHA256
