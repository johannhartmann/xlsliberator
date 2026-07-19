"""Office-free contract tests for the generic Docker application boundary."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest

from xlsliberator.application_showcase import (
    GUI_IMAGE,
    _require_success,
    build_candidate,
    bundle_application_replays,
    run_application_scenario,
)


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


def test_build_candidate_uses_only_the_pinned_worker_boundary(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    source = tmp_path / "source.xlsb"
    candidate = tmp_path / "candidate.zip"
    output = tmp_path / "target.ods"
    source.write_bytes(b"source")
    candidate.write_bytes(b"candidate")
    monkeypatch.setattr(
        "xlsliberator.application_showcase.LibreOfficeDockerRuntime",
        _FakeRuntime,
    )
    _FakeRuntime.instances.clear()

    result = build_candidate(source, candidate, output)

    assert result["status"] == "passed"
    runtime = _FakeRuntime.instances[0]
    assert runtime.image is None
    assert runtime.payload == {
        "op": "build_application_target",
        "input_path": str(source),
        "candidate_path": str(candidate),
        "output_path": str(output),
        "timeout_seconds": 120,
    }


def test_application_scenario_forwards_opaque_candidate_configuration(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    target = tmp_path / "target.ods"
    candidate = tmp_path / "candidate.zip"
    output = tmp_path / "evidence.zip"
    target.write_bytes(b"target")
    candidate.write_bytes(b"candidate")
    actions = [{"kind": "observe", "sheet": "Input", "address": "C2"}]
    monkeypatch.setattr(
        "xlsliberator.application_showcase.LibreOfficeDockerRuntime",
        _FakeRuntime,
    )
    _FakeRuntime.instances.clear()

    run_application_scenario(
        target,
        candidate,
        output,
        actions,
        scenario_id="boundary-case",
        adapter_config={"clock_enabled": False},
    )

    runtime = _FakeRuntime.instances[0]
    assert runtime.image == GUI_IMAGE
    assert runtime.payload is not None
    assert runtime.payload["op"] == "run_application_scenario"
    assert runtime.payload["candidate_path"] == str(candidate)
    assert runtime.payload["scenario_id"] == "boundary-case"
    assert runtime.payload["adapter_config"] == {"clock_enabled": False}
    assert runtime.payload["actions"] == actions


def test_replay_bundle_accepts_any_declared_safe_scenario_set(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    evidence: dict[str, Path] = {}
    for scenario_id in ("first", "boundary"):
        path = tmp_path / f"{scenario_id}.zip"
        path.write_bytes(scenario_id.encode())
        evidence[scenario_id] = path
    output = tmp_path / "replay.zip"
    monkeypatch.setattr(
        "xlsliberator.application_showcase.LibreOfficeDockerRuntime",
        _FakeRuntime,
    )
    _FakeRuntime.instances.clear()

    bundle_application_replays(evidence, output, replay_id="migration-42")

    runtime = _FakeRuntime.instances[0]
    assert runtime.image == GUI_IMAGE
    assert runtime.payload is not None
    assert runtime.payload["op"] == "bundle_application_replays"
    assert runtime.payload["scenario_ids"] == ["first", "boundary"]
    assert runtime.payload["replay_id"] == "migration-42"
    assert not Path(runtime.payload["input_path"]).exists()


def test_failed_scenario_preserves_worker_and_container_diagnostics() -> None:
    with pytest.raises(
        RuntimeError,
        match=r"(?s)bridge disposed.*office_exit_code=134.*container stderr:.*outer stderr",
    ):
        _require_success(
            {
                "success": False,
                "data": {"container_stderr": "outer stderr"},
                "error": {
                    "message": "bridge disposed",
                    "traceback": "office_exit_code=134",
                },
            },
            "migration candidate GUI scenario",
        )
