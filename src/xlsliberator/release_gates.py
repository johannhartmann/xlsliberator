"""Content-bound attestations for blocking Docker release gates."""

from __future__ import annotations

import hashlib
from datetime import datetime
from pathlib import Path
from typing import Literal

from pydantic import BaseModel, ConfigDict, Field

GateCheck = Literal["quality", "security", "office", "docker-web", "package"]

_RELEASE_INPUT_DIRECTORIES = (
    ".github",
    "docker",
    "office",
    "src",
    "tests",
    "tools",
)
_RELEASE_INPUT_FILES = (
    ".env.example",
    ".gitignore",
    "Dockerfile",
    "Makefile",
    "docker-compose.yml",
    "pyproject.toml",
    "sitecustomize.py",
)
_IGNORED_PARTS = frozenset({"__pycache__", ".mypy_cache", ".pytest_cache", ".ruff_cache"})


class GateAttestation(BaseModel):
    """Successful result emitted by one Docker-contained blocking gate."""

    model_config = ConfigDict(extra="forbid")

    schema_version: Literal["1.0.0"] = "1.0.0"
    check: GateCheck
    status: Literal["passed"]
    runner: Literal["xlsliberator-test-container"]
    workspace_sha256: str = Field(pattern=r"^[0-9a-f]{64}$")
    recorded_at: datetime


def release_workspace_sha256(root: Path) -> str:
    """Hash every release-relevant path while excluding generated caches/evidence."""
    files = [root / name for name in _RELEASE_INPUT_FILES]
    for directory_name in _RELEASE_INPUT_DIRECTORIES:
        directory = root / directory_name
        if directory.is_dir():
            files.extend(path for path in directory.rglob("*") if path.is_file())
    selected = sorted(
        {
            path
            for path in files
            if path.is_file()
            and not _IGNORED_PARTS.intersection(path.relative_to(root).parts)
            and path.suffix not in {".pyc", ".pyo"}
        },
        key=lambda path: path.relative_to(root).as_posix(),
    )
    digest = hashlib.sha256()
    for path in selected:
        relative = path.relative_to(root).as_posix().encode("utf-8")
        digest.update(len(relative).to_bytes(8, "big"))
        digest.update(relative)
        content = path.read_bytes()
        digest.update(len(content).to_bytes(8, "big"))
        digest.update(content)
    return digest.hexdigest()


def load_gate_attestation(
    path: Path, *, expected_check: Literal["quality", "security"], workspace_sha256: str
) -> GateAttestation:
    """Validate that an attestation belongs to this gate and exact workspace."""
    if not path.is_file():
        raise RuntimeError(f"required Docker {expected_check} attestation is absent: {path}")
    attestation = GateAttestation.model_validate_json(path.read_text(encoding="utf-8"))
    if attestation.check != expected_check:
        raise RuntimeError(
            f"expected {expected_check} attestation, received {attestation.check}: {path}"
        )
    if attestation.workspace_sha256 != workspace_sha256:
        raise RuntimeError(f"stale {expected_check} attestation does not match this workspace")
    return attestation
