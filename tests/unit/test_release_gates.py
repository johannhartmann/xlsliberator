"""Tests for content-bound Docker release-gate attestations."""

from __future__ import annotations

from datetime import UTC, datetime
from pathlib import Path

import pytest

from xlsliberator.release_gates import (
    GateAttestation,
    load_gate_attestation,
    release_workspace_sha256,
)


def test_workspace_digest_changes_with_release_input_but_ignores_cache(tmp_path: Path) -> None:
    source = tmp_path / "src"
    source.mkdir()
    module = source / "module.py"
    module.write_text("VALUE = 1\n", encoding="utf-8")
    first = release_workspace_sha256(tmp_path)

    cache = source / "__pycache__"
    cache.mkdir()
    (cache / "module.pyc").write_bytes(b"generated")
    assert release_workspace_sha256(tmp_path) == first

    module.write_text("VALUE = 2\n", encoding="utf-8")
    assert release_workspace_sha256(tmp_path) != first


def test_attestation_must_match_gate_and_exact_workspace(tmp_path: Path) -> None:
    workspace_sha256 = "a" * 64
    path = tmp_path / "quality.json"
    attestation = GateAttestation(
        check="quality",
        status="passed",
        runner="xlsliberator-test-container",
        workspace_sha256=workspace_sha256,
        recorded_at=datetime.now(UTC),
    )
    path.write_text(attestation.model_dump_json(), encoding="utf-8")

    assert (
        load_gate_attestation(
            path, expected_check="quality", workspace_sha256=workspace_sha256
        ).status
        == "passed"
    )
    with pytest.raises(RuntimeError, match="stale quality attestation"):
        load_gate_attestation(path, expected_check="quality", workspace_sha256="b" * 64)
    with pytest.raises(RuntimeError, match="expected security attestation"):
        load_gate_attestation(path, expected_check="security", workspace_sha256=workspace_sha256)


def test_missing_attestation_is_explicitly_unavailable(tmp_path: Path) -> None:
    with pytest.raises(RuntimeError, match="required Docker quality attestation is absent"):
        load_gate_attestation(
            tmp_path / "missing.json",
            expected_check="quality",
            workspace_sha256="a" * 64,
        )
