"""Authenticated workbook API served by the embedded Open-SWE LangGraph runtime."""

from __future__ import annotations

import base64
import binascii
import contextlib
import hashlib
import hmac
import json
import mimetypes
import os
import re
import shutil
import uuid
import zipfile
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

from fastapi import APIRouter, FastAPI, Header, HTTPException, Request
from fastapi.responses import Response
from langgraph_sdk import get_client

from xlsliberator.migration_check import load_acceptance
from xlsliberator.open_swe_agent.state import (
    append_event,
    read_state,
    thread_root,
    update_state,
    workspace_root,
    write_state,
)
from xlsliberator.scenarios.acceptance_evidence import inspect_acceptance_evidence
from xlsliberator.scenarios.models import MutationCampaign, ObservationKind
from xlsliberator.validation_models import GateExecutionStatus
from xlsliberator.xlsprobe import probe_workbook

_ARTIFACT_ID = re.compile(r"^[0-9a-f]{24}$")
_TERMINAL_RUN_STATES = frozenset({"error", "failed", "timeout", "interrupted"})
_MAX_SOURCE_BYTES = 64 * 1024 * 1024

router = APIRouter(prefix="/api/xlsliberator/migrations", tags=["xlsliberator"])
app = FastAPI(title="XLSLiberator Open-SWE", version="0.1.0")


def _client() -> Any:
    return get_client(url=os.environ.get("LANGGRAPH_URL", "http://127.0.0.1:2024"))


def _service_token() -> str:
    return os.environ.get("XLSLIBERATOR_OPEN_SWE_SERVICE_TOKEN", "")


def _authenticate(authorization: str | None, owner_header: str | None) -> str:
    expected = _service_token()
    supplied = authorization.removeprefix("Bearer ").strip() if authorization else ""
    owner = owner_header.strip() if owner_header else ""
    if not expected or not supplied or not hmac.compare_digest(supplied, expected) or not owner:
        raise HTTPException(401, "invalid migration service credentials")
    if len(owner) > 200 or any(ord(character) < 32 for character in owner):
        raise HTTPException(400, "invalid migration owner")
    return owner


def _safe_filename(value: object) -> str:
    if not isinstance(value, str):
        raise HTTPException(422, "artifact filename is required")
    plain = Path(value).name
    if plain != value or not plain or "\x00" in value:
        raise HTTPException(422, "invalid artifact filename")
    return re.sub(r"[^A-Za-z0-9._-]+", "-", plain)[:200]


def _model_configured() -> bool:
    return bool(os.environ.get("XLSLIBERATOR_OPEN_SWE_MODEL", "").strip())


def _reviewer_configured() -> bool:
    reviewer = os.environ.get("XLSLIBERATOR_REVIEWER_MODEL", "").strip()
    migration = os.environ.get("XLSLIBERATOR_OPEN_SWE_MODEL", "").strip()
    return bool(reviewer and reviewer != migration)


@app.get("/health")
async def health() -> dict[str, Any]:
    return {
        "status": "healthy",
        "runtime": "open-swe",
        "upstream_commit": os.environ.get("OPEN_SWE_UPSTREAM_COMMIT", "unknown"),
        "model_configured": _model_configured(),
        "reviewer_configured": _reviewer_configured(),
        "github_models_enabled": os.environ.get("XLSLIBERATOR_GITHUB_MODELS_ENABLED") == "1",
        "target_libreoffice_version": "26.2.4.2",
    }


@router.post("")
async def create_migration(
    request: Request,
    authorization: str | None = Header(default=None),
    x_xlsliberator_owner: str | None = Header(default=None),
) -> Response:
    owner = _authenticate(authorization, x_xlsliberator_owner)
    if not _model_configured():
        raise HTTPException(
            503,
            "XLSLIBERATOR_OPEN_SWE_MODEL must be configured explicitly; "
            "no paid model is selected automatically",
        )
    if not _reviewer_configured():
        raise HTTPException(
            503,
            "XLSLIBERATOR_REVIEWER_MODEL must be configured and differ from "
            "XLSLIBERATOR_OPEN_SWE_MODEL",
        )
    payload = await request.json()
    if not isinstance(payload, dict) or payload.get("owner_id") != owner:
        raise HTTPException(403, "migration owner mismatch")
    if payload.get("target_libreoffice_version") != "26.2.4.2":
        raise HTTPException(422, "only LibreOffice 26.2.4.2 is supported")
    artifact = payload.get("artifact")
    if not isinstance(artifact, dict):
        raise HTTPException(422, "workbook artifact is required")
    filename = _safe_filename(artifact.get("original_filename"))
    encoded = artifact.get("artifact_base64")
    if not isinstance(encoded, str):
        raise HTTPException(422, "workbook content is required")
    try:
        content = base64.b64decode(encoded, validate=True)
    except binascii.Error as exc:
        raise HTTPException(422, "invalid workbook encoding") from exc
    if not content or len(content) > _MAX_SOURCE_BYTES:
        raise HTTPException(413, "workbook exceeds the 64 MiB service limit")
    digest = hashlib.sha256(content).hexdigest()
    if artifact.get("sha256") != digest:
        raise HTTPException(422, "workbook digest mismatch")

    thread_id = str(uuid.uuid4())
    root = thread_root(thread_id)
    source_dir = root / "source"
    (root / "deliverables").mkdir(parents=True)
    (root / "evidence").mkdir()
    source_dir.mkdir()
    source = source_dir / filename
    source.write_bytes(content)
    now = datetime.now(UTC).isoformat()
    user_requirements = str(payload.get("user_requirements") or "")[:20_000]
    state: dict[str, Any] = {
        "thread_id": thread_id,
        "run_id": None,
        "owner_id": owner,
        "status": "queued",
        "source": f"source/{filename}",
        "source_sha256": digest,
        "user_requirements": user_requirements,
        "artifacts": [],
        "events": [],
        "created_at": now,
        "updated_at": now,
    }
    write_state(thread_id, state)
    append_event(thread_id, stage="upload", message="Private workbook accepted")

    client = _client()
    thread_created = False
    try:
        await client.threads.create(
            thread_id=thread_id,
            metadata={"owner_id": owner, "source": "xlsliberator-web"},
            if_exists="raise",
        )
        thread_created = True
        run = await client.runs.create(
            thread_id,
            "xlsliberator",
            input={
                "messages": [
                    {
                        "role": "user",
                        "content": _migration_prompt(filename, user_requirements),
                    }
                ]
            },
            config={"configurable": {"thread_id": thread_id}},
            multitask_strategy="interrupt",
            durability="sync",
            if_not_exists="create",
        )
    except Exception as exc:
        if thread_created:
            with contextlib.suppress(Exception):
                await client.threads.delete(thread_id)
        shutil.rmtree(root, ignore_errors=True)
        raise HTTPException(503, "Open-SWE runtime could not start migration") from exc
    run_id = run.get("run_id") if isinstance(run, dict) else None
    update_state(thread_id, run_id=run_id, status="running")
    response: dict[str, Any] = {
        "thread_id": thread_id,
        "run_id": run_id,
        "duplicate": False,
        "artifact_locations": {},
    }
    return Response(
        content=json.dumps(response),
        status_code=202,
        media_type="application/json",
    )


@router.get("/{thread_id}")
async def migration_status(
    thread_id: str,
    authorization: str | None = Header(default=None),
    x_xlsliberator_owner: str | None = Header(default=None),
) -> dict[str, Any]:
    owner = _authenticate(authorization, x_xlsliberator_owner)
    state = _owned_state(thread_id, owner)
    state = await _refresh_run_state(state)
    return _public_state(state)


@router.get("/{thread_id}/events")
async def migration_events(
    thread_id: str,
    since: int = 0,
    authorization: str | None = Header(default=None),
    x_xlsliberator_owner: str | None = Header(default=None),
) -> dict[str, Any]:
    owner = _authenticate(authorization, x_xlsliberator_owner)
    state = _owned_state(thread_id, owner)
    events = state.get("events")
    public_events = events if isinstance(events, list) else []
    start = max(0, since)
    return {
        "thread_id": thread_id,
        "events": public_events[start:],
        "next": len(public_events),
    }


@router.post("/{thread_id}/follow-ups")
async def follow_up(
    thread_id: str,
    request: Request,
    authorization: str | None = Header(default=None),
    x_xlsliberator_owner: str | None = Header(default=None),
) -> dict[str, Any]:
    owner = _authenticate(authorization, x_xlsliberator_owner)
    _owned_state(thread_id, owner)
    payload = await request.json()
    if not isinstance(payload, dict):
        raise HTTPException(422, "follow-up must be an object")
    requirements = str(payload.get("requirements") or "")[:20_000]
    dependency = payload.get("dependency")
    dependency_note = ""
    if isinstance(dependency, dict):
        filename = _safe_filename(dependency.get("original_filename"))
        encoded = dependency.get("artifact_base64")
        if not isinstance(encoded, str):
            raise HTTPException(422, "dependency content is required")
        try:
            content = base64.b64decode(encoded, validate=True)
        except binascii.Error as exc:
            raise HTTPException(422, "invalid dependency encoding") from exc
        if not content or len(content) > _MAX_SOURCE_BYTES:
            raise HTTPException(413, "dependency exceeds the 64 MiB service limit")
        digest = hashlib.sha256(content).hexdigest()
        if dependency.get("sha256") != digest:
            raise HTTPException(422, "dependency digest mismatch")
        destination = thread_root(thread_id) / "dependencies" / filename
        destination.parent.mkdir(exist_ok=True)
        destination.write_bytes(content)
        dependency_note = f"\nNew dependency: /workspace/dependencies/{filename}"
    if not requirements and not dependency_note:
        raise HTTPException(422, "follow-up text or dependency is required")
    run = await _client().runs.create(
        thread_id,
        "xlsliberator",
        input={"messages": [{"role": "user", "content": requirements + dependency_note}]},
        config={"configurable": {"thread_id": thread_id}},
        multitask_strategy="interrupt",
        durability="sync",
        if_not_exists="create",
    )
    run_id = run.get("run_id") if isinstance(run, dict) else None
    update_state(thread_id, run_id=run_id, status="running", artifacts=[])
    append_event(thread_id, stage="follow_up", message="Follow-up added to Open-SWE thread")
    return {"thread_id": thread_id, "run_id": run_id}


@router.post("/{thread_id}/cancel")
async def cancel_migration(
    thread_id: str,
    authorization: str | None = Header(default=None),
    x_xlsliberator_owner: str | None = Header(default=None),
) -> dict[str, Any]:
    owner = _authenticate(authorization, x_xlsliberator_owner)
    state = _owned_state(thread_id, owner)
    reviewer_run_id = state.get("reviewer_run_id")
    reviewer_thread_id = state.get("reviewer_thread_id")
    run_id = reviewer_run_id or state.get("run_id")
    active_thread_id = (
        reviewer_thread_id
        if isinstance(reviewer_run_id, str) and isinstance(reviewer_thread_id, str)
        else thread_id
    )
    if isinstance(run_id, str) and run_id:
        await _client().runs.cancel(active_thread_id, run_id, wait=False)
    update_state(thread_id, status="cancelled")
    append_event(thread_id, stage="final", message="Migration cancelled", status="cancelled")
    return {"thread_id": thread_id, "status": "cancelled"}


@router.delete("/{thread_id}")
async def delete_migration(
    thread_id: str,
    authorization: str | None = Header(default=None),
    x_xlsliberator_owner: str | None = Header(default=None),
) -> dict[str, Any]:
    owner = _authenticate(authorization, x_xlsliberator_owner)
    state = _owned_state(thread_id, owner)
    review_id = state.get("review_id")
    reviewer_run_id = state.get("reviewer_run_id")
    reviewer_thread_id = state.get("reviewer_thread_id")
    run_id = reviewer_run_id or state.get("run_id")
    active_thread_id = (
        reviewer_thread_id
        if isinstance(reviewer_run_id, str) and isinstance(reviewer_thread_id, str)
        else thread_id
    )
    if isinstance(run_id, str) and run_id:
        with contextlib.suppress(Exception):
            await _client().runs.cancel(active_thread_id, run_id, wait=False)
    try:
        await _client().threads.delete(thread_id)
    finally:
        if isinstance(reviewer_thread_id, str):
            with contextlib.suppress(Exception):
                await _client().threads.delete(reviewer_thread_id)
        shutil.rmtree(thread_root(thread_id), ignore_errors=True)
        if isinstance(review_id, str):
            shutil.rmtree(workspace_root() / ".reviews" / review_id, ignore_errors=True)
    return {"thread_id": thread_id, "status": "cleaned"}


@router.get("/{thread_id}/artifacts/{artifact_id}")
async def download_artifact(
    thread_id: str,
    artifact_id: str,
    authorization: str | None = Header(default=None),
    x_xlsliberator_owner: str | None = Header(default=None),
) -> Response:
    owner = _authenticate(authorization, x_xlsliberator_owner)
    state = _owned_state(thread_id, owner)
    if _ARTIFACT_ID.fullmatch(artifact_id) is None:
        raise HTTPException(404, "artifact not found")
    for item in state.get("artifacts", []):
        if isinstance(item, dict) and item.get("id") == artifact_id:
            relative = item.get("path")
            if not isinstance(relative, str):
                break
            destination = (thread_root(thread_id) / relative).resolve()
            try:
                destination.relative_to(thread_root(thread_id).resolve())
            except ValueError:
                break
            if destination.is_file():
                return Response(
                    content=destination.read_bytes(),
                    media_type=str(item.get("media_type") or "application/octet-stream"),
                )
    raise HTTPException(404, "artifact not found")


def _owned_state(thread_id: str, owner: str) -> dict[str, Any]:
    try:
        parsed = str(uuid.UUID(thread_id))
    except ValueError as exc:
        raise HTTPException(404, "migration not found") from exc
    if parsed != thread_id:
        raise HTTPException(404, "migration not found")
    try:
        state = read_state(thread_id)
    except (OSError, ValueError, json.JSONDecodeError) as exc:
        raise HTTPException(404, "migration not found") from exc
    if state.get("owner_id") != owner:
        raise HTTPException(404, "migration not found")
    return state


async def _refresh_run_state(state: dict[str, Any]) -> dict[str, Any]:
    if state.get("status") in {"complete", "failed", "cancelled", "rejected"}:
        return state
    thread_id = str(state["thread_id"])
    active_run_id = state.get("reviewer_run_id") or state.get("run_id")
    if not isinstance(active_run_id, str) or not active_run_id:
        return state
    reviewer_thread_id = state.get("reviewer_thread_id")
    active_thread_id = (
        reviewer_thread_id
        if state.get("reviewer_run_id") and isinstance(reviewer_thread_id, str)
        else thread_id
    )
    run = await _client().runs.get(active_thread_id, active_run_id)
    run_status = run.get("status") if isinstance(run, dict) else getattr(run, "status", None)
    if run_status == "success":
        if state.get("reviewer_run_id"):
            return _finalize_reviewer_success(thread_id, active_run_id)
        return await _start_reviewer(thread_id)
    if run_status in _TERMINAL_RUN_STATES:
        if state.get("reviewer_run_id"):
            terminal = "rejected"
        else:
            terminal = "cancelled" if run_status == "interrupted" else "failed"
        updated = update_state(thread_id, status=terminal)
        append_event(
            thread_id,
            stage="final",
            message=f"Open-SWE run ended as {run_status}",
            status=terminal,
        )
        return updated
    expected_status = "reviewing" if state.get("reviewer_run_id") else "running"
    if run_status in {"pending", "running"} and state.get("status") != expected_status:
        return update_state(thread_id, status=expected_status)
    return state


async def _start_reviewer(thread_id: str) -> dict[str, Any]:
    """Snapshot passed migration evidence and start the separate reviewer graph."""
    root = thread_root(thread_id)
    state = read_state(thread_id)
    source = root / str(state["source"])
    target = root / "deliverables" / "target.ods"
    reviewer_model = os.environ.get("XLSLIBERATOR_REVIEWER_MODEL", "").strip()
    migration_model = os.environ.get("XLSLIBERATOR_OPEN_SWE_MODEL", "").strip()
    if not reviewer_model or reviewer_model == migration_model:
        append_event(
            thread_id,
            stage="reviewer",
            message="Independent reviewer model is missing or not distinct",
            status="rejected",
        )
        return update_state(thread_id, status="rejected")
    if not _migration_evidence_complete(root, source, target, state):
        append_event(
            thread_id,
            stage="reviewer",
            message="Review was not started because deterministic evidence is incomplete",
            status="rejected",
        )
        return update_state(thread_id, status="rejected")

    review_id = str(uuid.uuid4())
    reviewer_thread_id = str(uuid.uuid4())
    snapshot = workspace_root() / ".reviews" / review_id
    reviewer_thread_created = False
    try:
        _create_review_snapshot(root, source, snapshot, review_id)
        await _client().threads.create(
            thread_id=reviewer_thread_id,
            metadata={"parent_thread_id": thread_id, "review_id": review_id},
            if_exists="raise",
        )
        reviewer_thread_created = True
        run = await _client().runs.create(
            reviewer_thread_id,
            "xlsliberator_reviewer",
            input={
                "messages": [
                    {
                        "role": "user",
                        "content": "Review the immutable migration snapshot and submit one decision.",
                    }
                ]
            },
            config={"configurable": {"thread_id": reviewer_thread_id, "review_id": review_id}},
            multitask_strategy="interrupt",
            durability="sync",
            if_not_exists="create",
        )
        reviewer_run_id = run.get("run_id") if isinstance(run, dict) else None
        if not isinstance(reviewer_run_id, str) or not reviewer_run_id:
            raise ValueError("reviewer graph did not return a run ID")
    except Exception:
        if reviewer_thread_created:
            with contextlib.suppress(Exception):
                await _client().threads.delete(reviewer_thread_id)
        shutil.rmtree(snapshot, ignore_errors=True)
        append_event(
            thread_id,
            stage="reviewer",
            message="Independent reviewer graph could not be started",
            status="rejected",
        )
        return update_state(thread_id, status="rejected")
    append_event(thread_id, stage="reviewer", message="Independent reviewer graph started")
    return update_state(
        thread_id,
        status="reviewing",
        review_id=review_id,
        reviewer_thread_id=reviewer_thread_id,
        reviewer_run_id=reviewer_run_id,
    )


def _create_review_snapshot(root: Path, source: Path, snapshot: Path, review_id: str) -> None:
    """Copy only review inputs outside the migration graph's filesystem root."""
    for path in root.rglob("*"):
        if path.is_symlink():
            raise ValueError(f"review input contains a symlink: {path}")
    snapshot.mkdir(parents=True)
    (snapshot / "source").mkdir()
    shutil.copy2(source, snapshot / "source" / source.name)
    for directory in ("dossier", "deliverables", "evidence"):
        source_directory = root / directory
        if source_directory.is_dir():
            shutil.copytree(
                source_directory,
                snapshot / directory,
                ignore=(
                    shutil.ignore_patterns("reviewer.json") if directory == "evidence" else None
                ),
            )
    acceptance = root / "acceptance.yaml"
    shutil.copy2(acceptance, snapshot / acceptance.name)
    test_files = [
        snapshot / "acceptance.yaml",
        snapshot / "evidence" / "public-acceptance" / "manifest.json",
    ]
    mutation = snapshot / "evidence" / "mutations" / "mutation-report.json"
    if mutation.is_file():
        test_files.append(mutation)
    test_set_identity = hashlib.sha256(
        "".join(_hash_file(path) for path in test_files).encode()
    ).hexdigest()
    context = {
        "schema_version": "1.0.0",
        "review_id": review_id,
        "source_path": f"source/{source.name}",
        "target_path": "deliverables/target.ods",
        "source_sha256": _hash_file(snapshot / "source" / source.name),
        "candidate_sha256": _hash_file(snapshot / "deliverables" / "target.ods"),
        "test_set_identity": test_set_identity,
    }
    (snapshot / "review-context.json").write_text(
        json.dumps(context, indent=2, sort_keys=True) + "\n",
        encoding="utf-8",
    )


def _finalize_reviewer_success(thread_id: str, reviewer_run_id: str) -> dict[str, Any]:
    """Validate the isolated reviewer record and bind it to its LangGraph run."""
    root = thread_root(thread_id)
    state = read_state(thread_id)
    review_id = state.get("review_id")
    if not isinstance(review_id, str):
        return update_state(thread_id, status="rejected")
    snapshot = workspace_root() / ".reviews" / review_id
    decision = _read_json(snapshot / "review-decision.json")
    context = _read_json(snapshot / "review-context.json")
    source = root / str(state["source"])
    target = root / "deliverables" / "target.ods"
    valid = (
        decision.get("review_id") == review_id
        and decision.get("decision") == "APPROVE"
        and decision.get("reviewer_model")
        == os.environ.get("XLSLIBERATOR_REVIEWER_MODEL", "").strip()
        and decision.get("source_sha256") == _hash_file(source)
        and decision.get("candidate_sha256") == _hash_file(target)
        and decision.get("test_set_identity") == context.get("test_set_identity")
    )
    if not valid:
        append_event(
            thread_id,
            stage="reviewer",
            message="Independent reviewer did not approve the content-bound candidate",
            status="rejected",
        )
        return update_state(thread_id, status="rejected")
    record = {
        **decision,
        "reviewer_graph": "xlsliberator_reviewer",
        "reviewer_thread_id": state.get("reviewer_thread_id"),
        "reviewer_run_id": reviewer_run_id,
    }
    destination = root / "evidence" / "reviewer.json"
    temporary = destination.with_suffix(".json.tmp")
    temporary.write_text(json.dumps(record, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    temporary.replace(destination)
    return _finalize_success(thread_id)


def _finalize_success(thread_id: str) -> dict[str, Any]:
    root = thread_root(thread_id)
    state = read_state(thread_id)
    source = root / str(state["source"])
    target = root / "deliverables" / "target.ods"
    reviewer = _read_json(root / "evidence" / "reviewer.json")
    required = _migration_evidence_complete(root, source, target, state) and (
        reviewer.get("decision") == "APPROVE"
        and reviewer.get("reviewer_run_id") == state.get("reviewer_run_id")
        and reviewer.get("reviewer_thread_id") == state.get("reviewer_thread_id")
        and reviewer.get("reviewer_graph") == "xlsliberator_reviewer"
    )
    if not required:
        updated = update_state(thread_id, status="rejected", artifacts=[])
        append_event(
            thread_id,
            stage="reviewer",
            message="Completion rejected because required deterministic evidence is missing",
            status="rejected",
        )
        return updated
    artifacts = _catalog_artifacts(root)
    updated = update_state(thread_id, status="complete", artifacts=artifacts)
    append_event(
        thread_id,
        stage="final",
        message="Open-SWE migration approved with deterministic evidence",
        status="complete",
    )
    return updated


def _migration_evidence_complete(
    root: Path,
    source: Path,
    target: Path,
    state: dict[str, Any],
) -> bool:
    """Validate deterministic evidence before and after independent review."""
    try:
        return (
            source.is_file()
            and _hash_file(source) == state.get("source_sha256")
            and target.is_file()
            and zipfile.is_zipfile(target)
            and (root / "deliverables" / "report.json").is_file()
            and (root / "deliverables" / "report.md").is_file()
            and _read_json(root / "evidence" / "certification.json").get("operation_status")
            == "passed"
            and _read_json(root / "evidence" / "save-reopen.json").get("operation_status")
            == "passed"
            and _behavioral_evidence_complete(root, source, target)
        )
    except (OSError, ValueError, json.JSONDecodeError, zipfile.BadZipFile):
        return False


def _behavioral_evidence_complete(root: Path, source: Path, target: Path) -> bool:
    """Require feature-appropriate, content-bound target behavior evidence."""
    report = probe_workbook(source)
    behavioral = bool(
        report.formulas or report.vba_modules or report.controls or report.dependencies
    )
    if not behavioral:
        return True

    acceptance_path = root / "acceptance.yaml"
    if not acceptance_path.is_file():
        return False
    acceptance = load_acceptance(acceptance_path)
    evidence = inspect_acceptance_evidence(root / "evidence" / "public-acceptance")
    if (
        evidence.status is not GateExecutionStatus.PASSED
        or evidence.workbook_sha256 != _hash_file(target)
        or _read_json(root / "evidence" / "public-acceptance.json").get("operation_status")
        != "passed"
    ):
        return False

    observations = {
        request.kind
        for step in acceptance.scenario.steps
        for request in (*step.observations_before, *step.observations_after)
    }
    if report.controls and not _gui_evidence_complete(root, target):
        return False
    if report.dependencies and not (
        {ObservationKind.MOCKED_CALLS, ObservationKind.RUNTIME_ERRORS} <= observations
    ):
        return False

    if report.formulas or report.vba_modules:
        campaign = MutationCampaign.model_validate_json(
            (root / "evidence" / "mutations" / "mutation-report.json").read_text(encoding="utf-8")
        )
        if (
            campaign.status is not GateExecutionStatus.PASSED
            or campaign.source_workbook_sha256 != _hash_file(target)
            or _read_json(root / "evidence" / "mutation-suite.json").get("operation_status")
            != "passed"
        ):
            return False
    return True


def _gui_evidence_complete(root: Path, target: Path) -> bool:
    """Validate one server-run GUI archive bound to the target and candidate package."""
    candidate = _read_json(root / "evidence" / "candidate-package.json")
    candidate_sha256 = candidate.get("candidate_bundle_sha256")
    if not isinstance(candidate_sha256, str) or not candidate_sha256:
        return False
    for wrapper_path in sorted((root / "evidence").glob("gui-*.json")):
        wrapper = _read_json(wrapper_path)
        data = wrapper.get("data")
        if wrapper.get("operation_status") != "passed" or not isinstance(data, dict):
            continue
        scenario_id = data.get("scenario_id")
        if (
            not isinstance(scenario_id, str)
            or re.fullmatch(r"[A-Za-z0-9_.-]{1,100}", scenario_id) is None
        ):
            continue
        archive_path = root / "evidence" / "gui" / f"{scenario_id}.zip"
        try:
            with zipfile.ZipFile(archive_path) as archive:
                member_names = archive.namelist()
                names = set(member_names)
                if (
                    "result.json" not in names
                    or any(name.startswith("/") or ".." in Path(name).parts for name in names)
                    or len(names) != len(member_names)
                ):
                    continue
                payload = archive.read("result.json")
                if len(payload) > 2 * 1024**2:
                    continue
                result = json.loads(payload)
        except (OSError, KeyError, UnicodeDecodeError, json.JSONDecodeError, zipfile.BadZipFile):
            continue
        if not isinstance(result, dict):
            continue
        operations = result.get("operations")
        kinds = {
            operation.get("kind")
            for operation in operations or []
            if isinstance(operation, dict) and operation.get("status") == "passed"
        }
        screenshots = result.get("screenshots")
        if (
            result.get("status") == "passed"
            and result.get("scenario_id") == scenario_id
            and result.get("target_sha256") == _hash_file(target)
            and result.get("candidate_bundle_sha256") == candidate_sha256
            and "click_control" in kinds
            and "screenshot" in kinds
            and isinstance(screenshots, list)
            and screenshots
            and all(isinstance(name, str) and name in names for name in screenshots)
        ):
            return True
    return False


def _hash_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _catalog_artifacts(root: Path) -> list[dict[str, Any]]:
    artifacts: list[dict[str, Any]] = []
    for directory in (root / "deliverables", root / "evidence"):
        for path in sorted(directory.rglob("*")):
            if not path.is_file():
                continue
            content = path.read_bytes()
            relative = str(path.relative_to(root))
            name = path.name
            kind = (
                "ods"
                if name == "target.ods"
                else ("report" if name.startswith("report.") else "evidence")
            )
            media_type = (
                "application/vnd.oasis.opendocument.spreadsheet"
                if name == "target.ods"
                else mimetypes.guess_type(name)[0] or "application/octet-stream"
            )
            artifacts.append(
                {
                    "id": hashlib.sha256(content).hexdigest()[:24],
                    "name": name,
                    "kind": kind,
                    "media_type": media_type,
                    "size": len(content),
                    "path": relative,
                }
            )
    return artifacts


def _read_json(path: Path) -> dict[str, Any]:
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}
    return payload if isinstance(payload, dict) else {}


def _public_state(state: dict[str, Any]) -> dict[str, Any]:
    return {
        "thread_id": state["thread_id"],
        "run_id": state.get("run_id"),
        "status": state.get("status"),
        "artifacts": [
            {key: value for key, value in item.items() if key != "path"}
            for item in state.get("artifacts", [])
            if isinstance(item, dict)
        ],
    }


def _migration_prompt(filename: str, requirements: str) -> str:
    extra = requirements.strip() or "Preserve all discoverable workbook behavior."
    return (
        f"Migrate /workspace/source/{filename} to LibreOffice 26.2.4.2.\n"
        f"User requirements: {extra}\n"
        "Produce the complete dossier, target-native implementation, deterministic "
        "evidence, independent review, and required deliverables."
    )


app.include_router(router)
