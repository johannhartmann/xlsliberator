"""Transactional machine-readable and Markdown acceptance evidence."""

from __future__ import annotations

import hashlib
import json
import os
import secrets
import shutil
import tempfile
from datetime import UTC, datetime
from pathlib import Path, PurePosixPath
from uuid import uuid4

from pydantic import BaseModel

from xlsliberator.scenarios.models import (
    AcceptanceDefinition,
    AcceptanceEvaluation,
    AcceptanceEvidenceManifest,
    RuntimeTrace,
)


def write_acceptance_evidence(
    directory: Path,
    *,
    workbook: Path,
    acceptance: AcceptanceDefinition,
    trace: RuntimeTrace,
    evaluation: AcceptanceEvaluation,
) -> AcceptanceEvidenceManifest:
    """Atomically create one complete, content-addressed evidence directory."""
    destination = directory.resolve()
    if destination.exists():
        raise FileExistsError(f"evidence directory already exists: {destination}")
    destination.parent.mkdir(parents=True, exist_ok=True)
    temporary = Path(tempfile.mkdtemp(prefix=f".{destination.name}.", dir=destination.parent))
    try:
        files = {
            "acceptance.json": acceptance.model_dump_json(indent=2) + "\n",
            "trace.json": trace.model_dump_json(indent=2) + "\n",
            "evaluation.json": evaluation.model_dump_json(indent=2) + "\n",
            "report.md": render_acceptance_report(acceptance, trace, evaluation),
        }
        for name, payload in files.items():
            _write_text(temporary / name, payload)
        file_hashes = {name: _hash_file(temporary / name) for name in sorted(files)}
        manifest = AcceptanceEvidenceManifest(
            evidence_id=uuid4().hex,
            created_at=datetime.now(UTC),
            migration_id=acceptance.migration.id,
            scenario_id=acceptance.scenario.id,
            status=evaluation.status,
            workbook=workbook.name,
            workbook_sha256=_hash_file(workbook),
            acceptance_definition="acceptance.json",
            execution_trace="trace.json",
            evaluation="evaluation.json",
            markdown_report="report.md",
            file_sha256=file_hashes,
        )
        _write_model(temporary / "manifest.json", manifest)
        _fsync_directory(temporary)
        os.replace(temporary, destination)
        _fsync_directory(destination.parent)
    except Exception:
        shutil.rmtree(temporary, ignore_errors=True)
        raise
    return manifest


def inspect_acceptance_evidence(directory: Path) -> AcceptanceEvidenceManifest:
    """Validate references, content hashes, and typed evidence documents."""
    root = directory.resolve()
    payload = json.loads((root / "manifest.json").read_text(encoding="utf-8"))
    manifest = AcceptanceEvidenceManifest.model_validate(payload)
    acceptance_path = _safe_reference(root, manifest.acceptance_definition)
    trace_path = _safe_reference(root, manifest.execution_trace)
    evaluation_path = _safe_reference(root, manifest.evaluation)
    acceptance = AcceptanceDefinition.model_validate_json(
        acceptance_path.read_text(encoding="utf-8")
    )
    trace = RuntimeTrace.model_validate_json(trace_path.read_text(encoding="utf-8"))
    evaluation = AcceptanceEvaluation.model_validate_json(
        evaluation_path.read_text(encoding="utf-8")
    )
    _safe_reference(root, manifest.markdown_report)
    required_references = {
        manifest.acceptance_definition,
        manifest.execution_trace,
        manifest.evaluation,
        manifest.markdown_report,
    }
    if set(manifest.file_sha256) != required_references:
        raise ValueError("evidence manifest hashes must cover every referenced evidence file")
    for reference, expected_hash in manifest.file_sha256.items():
        actual_hash = _hash_file(_safe_reference(root, reference))
        if not secrets.compare_digest(expected_hash, actual_hash):
            raise ValueError(
                f"evidence hash mismatch for {reference}: "
                f"expected {expected_hash}, found {actual_hash}"
            )
    if acceptance.migration.id != manifest.migration_id:
        raise ValueError("evidence manifest migration does not match its acceptance definition")
    if acceptance.scenario.id != manifest.scenario_id:
        raise ValueError("evidence manifest scenario does not match its acceptance definition")
    if trace.scenario_id != manifest.scenario_id:
        raise ValueError("evidence execution trace scenario does not match its manifest")
    if evaluation.migration_id != manifest.migration_id:
        raise ValueError("evidence evaluation migration does not match its manifest")
    if evaluation.scenario_id != manifest.scenario_id:
        raise ValueError("evidence evaluation scenario does not match its manifest")
    if evaluation.trace_id != trace.trace_id:
        raise ValueError("evidence evaluation trace does not match its execution trace")
    if evaluation.status is not manifest.status:
        raise ValueError("evidence manifest status does not match its evaluation")
    return manifest


def render_acceptance_report(
    acceptance: AcceptanceDefinition,
    trace: RuntimeTrace,
    evaluation: AcceptanceEvaluation,
) -> str:
    """Render a concise Markdown report from the typed evidence."""
    lines = [
        f"# Migration acceptance: {_markdown(acceptance.migration.title)}",
        "",
        f"- Migration: `{_markdown(acceptance.migration.id)}`",
        f"- Scenario: `{_markdown(acceptance.scenario.id)}`",
        f"- Trace: `{_markdown(trace.trace_id)}`",
        f"- Runtime: `{_markdown(trace.runtime_identity.runtime_kind)}`",
        f"- Result: **{evaluation.status.value.upper()}**",
        f"- Oracle policy: `{acceptance.migration.oracle_policy}`",
        "",
        "Cached Excel values are not used as an authoritative oracle. "
        "Expectations come from the authored and independently reviewed acceptance requirements.",
        "",
        "## Actions",
        "",
        "| Step | Status |",
        "| --- | --- |",
    ]
    lines.extend(
        f"| {_markdown(step_id)} | {status.value} |"
        for step_id, status in evaluation.action_statuses.items()
    )
    lines.extend(
        [
            "",
            "## Assertions",
            "",
            "| Step | Observation | Phase | Required | Status | Reason |",
            "| --- | --- | --- | --- | --- | --- |",
        ]
    )
    lines.extend(
        "| {step} | {observation} | {phase} | {required} | {status} | {reason} |".format(
            step=_markdown(item.step_id),
            observation=_markdown(item.observation_id),
            phase=item.phase,
            required="yes" if item.required else "no",
            status=item.status.value,
            reason=_markdown(item.reason or ""),
        )
        for item in evaluation.assertions
    )
    if evaluation.required_failures:
        lines.extend(["", "## Required failures", ""])
        lines.extend(f"- {_markdown(item)}" for item in evaluation.required_failures)
    return "\n".join(lines) + "\n"


def _safe_reference(root: Path, reference: str) -> Path:
    pure = PurePosixPath(reference)
    if pure.is_absolute() or ".." in pure.parts or not pure.parts:
        raise ValueError(f"unsafe evidence reference: {reference}")
    path = root.joinpath(*pure.parts)
    resolved = path.resolve()
    if path.is_symlink() or not resolved.is_relative_to(root):
        raise ValueError(f"unsafe evidence reference: {reference}")
    if not path.is_file():
        raise ValueError(f"evidence reference is missing: {reference}")
    return resolved


def _write_model(path: Path, model: BaseModel) -> None:
    _write_text(path, model.model_dump_json(indent=2) + "\n")


def _write_text(path: Path, payload: str) -> None:
    with path.open("x", encoding="utf-8") as handle:
        handle.write(payload)
        handle.flush()
        os.fsync(handle.fileno())


def _hash_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _fsync_directory(path: Path) -> None:
    descriptor = os.open(path, os.O_RDONLY)
    try:
        os.fsync(descriptor)
    finally:
        os.close(descriptor)


def _markdown(value: str) -> str:
    return value.replace("|", "\\|").replace("\n", " ")
