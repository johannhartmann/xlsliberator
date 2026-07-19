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

from xlsliberator.open_swe_agent.state import (
    append_event,
    read_state,
    thread_root,
    update_state,
    write_state,
)

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


@app.get("/health")
async def health() -> dict[str, Any]:
    return {
        "status": "healthy",
        "runtime": "open-swe",
        "upstream_commit": os.environ.get("OPEN_SWE_UPSTREAM_COMMIT", "unknown"),
        "model_configured": _model_configured(),
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
    run_id = state.get("run_id")
    if isinstance(run_id, str) and run_id:
        await _client().runs.cancel(thread_id, run_id, wait=False)
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
    run_id = state.get("run_id")
    if isinstance(run_id, str) and run_id:
        with contextlib.suppress(Exception):
            await _client().runs.cancel(thread_id, run_id, wait=False)
    try:
        await _client().threads.delete(thread_id)
    finally:
        shutil.rmtree(thread_root(thread_id), ignore_errors=True)
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
    runs = await _client().runs.list(thread_id, limit=1)
    if not runs:
        return state
    run = runs[0]
    run_status = run.get("status") if isinstance(run, dict) else getattr(run, "status", None)
    if run_status == "success":
        return _finalize_success(thread_id)
    if run_status in _TERMINAL_RUN_STATES:
        terminal = "cancelled" if run_status == "interrupted" else "failed"
        updated = update_state(thread_id, status=terminal)
        append_event(
            thread_id,
            stage="final",
            message=f"Open-SWE run ended as {run_status}",
            status=terminal,
        )
        return updated
    if run_status in {"pending", "running"} and state.get("status") != "running":
        return update_state(thread_id, status="running")
    return state


def _finalize_success(thread_id: str) -> dict[str, Any]:
    root = thread_root(thread_id)
    target = root / "deliverables" / "target.ods"
    certification = _read_json(root / "evidence" / "certification.json")
    save_reopen = _read_json(root / "evidence" / "save-reopen.json")
    reviewer = _read_json(root / "evidence" / "reviewer.json")
    required = (
        target.is_file()
        and zipfile.is_zipfile(target)
        and certification.get("operation_status") == "passed"
        and save_reopen.get("operation_status") == "passed"
        and reviewer.get("decision") == "APPROVE"
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
