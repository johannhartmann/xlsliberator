"""HTML and JSON routes for conversion jobs."""

from __future__ import annotations

import json
import math
import re
from pathlib import Path
from typing import Annotated, Any

from fastapi import APIRouter, BackgroundTasks, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, RedirectResponse, Response
from fastapi.templating import Jinja2Templates
from pydantic import BaseModel, ConfigDict, Field

from xlsliberator.validation_models import GateExecutionStatus
from xlsliberator.web.jobs import JobPhase, JobStore, WebJob, public_job_dict
from xlsliberator.web.open_swe import OpenSWEError
from xlsliberator.web.runner import WebJobRunner
from xlsliberator.web.schemas import WebSettings
from xlsliberator.web.security import (
    UploadValidationError,
    generate_job_id,
    safe_download_stem,
    safe_job_paths,
    validate_upload_filename,
    validate_upload_signature,
)

templates = Jinja2Templates(directory=str(Path(__file__).parent / "templates"))

GITHUB_URL = "https://github.com/johannhartmann/xlsliberator"


class FollowUpMessage(BaseModel):
    """Bounded message for an existing workbook thread."""

    model_config = ConfigDict(extra="forbid")

    message: str = Field(min_length=1, max_length=20_000)


def create_router(store: JobStore, runner: WebJobRunner, settings: WebSettings) -> APIRouter:
    """Create routes bound to the provided job store and runner."""
    router = APIRouter()

    @router.get("/", response_class=HTMLResponse)
    def index(request: Request) -> Response:
        return templates.TemplateResponse(
            request,
            "landing.html",
            {
                "max_upload_mb": settings.max_upload_mb,
                "github_url": GITHUB_URL,
                "demo_host": "127.0.0.1:8080",
            },
        )

    @router.post("/jobs")
    async def create_job(
        request: Request,
        background_tasks: BackgroundTasks,
        file: Annotated[UploadFile, File()],
    ) -> Response:
        return await _handle_upload(request, background_tasks, file, store, runner, settings)

    @router.post("/api/jobs")
    async def create_api_job(
        request: Request,
        background_tasks: BackgroundTasks,
        file: Annotated[UploadFile, File()],
    ) -> Response:
        return await _handle_upload(
            request,
            background_tasks,
            file,
            store,
            runner,
            settings,
            force_json=True,
        )

    @router.get("/jobs/{job_id}", response_class=HTMLResponse)
    def job_page(request: Request, job_id: str) -> Response:
        job = _get_job_or_404(store, job_id)
        return templates.TemplateResponse(
            request,
            "job.html",
            {
                "job": job,
                "report_summary": _load_report_summary(job),
                "public_job_json": json.dumps(public_job_dict(job)),
            },
        )

    @router.get("/jobs/{job_id}/showcase", response_class=HTMLResponse)
    def showcase_replay(request: Request, job_id: str) -> Response:
        job = _get_job_or_404(store, job_id)
        if job.status != JobPhase.COMPLETED or job.operation_status != GateExecutionStatus.PASSED:
            raise HTTPException(status_code=409, detail="Showcase evidence is not complete")
        replay = _load_showcase_replay(job)
        return templates.TemplateResponse(
            request,
            "showcase.html",
            {
                "job": job,
                "replay": replay,
            },
        )

    @router.get("/api/jobs/{job_id}")
    def api_job(job_id: str) -> dict[str, Any]:
        public = store.public_job(job_id)
        if public is None:
            raise HTTPException(status_code=404, detail="Unknown job")
        return public

    @router.get("/api/jobs/{job_id}/report")
    def api_job_report(job_id: str) -> dict[str, Any]:
        job = _get_job_or_404(store, job_id)
        summary = _load_report_summary(job)
        if summary is None:
            raise HTTPException(status_code=404, detail="Report not available")
        return summary

    @router.get("/api/jobs/{job_id}/events")
    def job_events(job_id: str, since: int = 0) -> dict[str, Any]:
        public = store.public_job(job_id)
        if public is None:
            raise HTTPException(status_code=404, detail="Unknown job")
        events = public["events"][max(0, since) :]
        return {"job_id": job_id, "events": events, "next": since + len(events)}

    @router.post("/api/jobs/{job_id}/cancel")
    def cancel_job(job_id: str) -> dict[str, Any]:
        try:
            job = store.request_cancel(job_id)
        except KeyError as exc:
            raise HTTPException(status_code=404, detail="Unknown job") from exc
        except ValueError as exc:
            raise HTTPException(status_code=409, detail=str(exc)) from exc
        try:
            runner.cancel(job_id)
        except OpenSWEError as exc:
            raise HTTPException(status_code=502, detail=str(exc)) from exc
        return store.public_job(job_id) or public_job_dict(job)

    @router.post("/jobs/{job_id}/cancel")
    def cancel_job_form(job_id: str) -> RedirectResponse:
        cancel_job(job_id)
        return RedirectResponse(f"/jobs/{job_id}", status_code=303)

    @router.post("/api/jobs/{job_id}/messages")
    def add_message(job_id: str, body: FollowUpMessage) -> dict[str, Any]:
        _get_job_or_404(store, job_id)
        try:
            runner.follow_up(job_id, requirements=body.message)
            runner.resume(job_id)
        except OpenSWEError as exc:
            raise HTTPException(status_code=409, detail=str(exc)) from exc
        return store.public_job(job_id) or {}

    @router.post("/jobs/{job_id}/follow-up")
    async def add_message_form(
        job_id: str,
        requirements: Annotated[str, Form()],
        file: Annotated[UploadFile | None, File()] = None,
    ) -> RedirectResponse:
        dependency = None
        media_type = "application/octet-stream"
        if file is not None and file.filename:
            dependency = await _write_dependency(file, store, job_id, settings.max_upload_mb)
            media_type = file.content_type or media_type
        if not requirements.strip() and dependency is None:
            raise HTTPException(status_code=400, detail="Message or dependency is required")
        try:
            runner.follow_up(
                job_id,
                requirements=requirements,
                dependency=dependency,
                media_type=media_type,
            )
            runner.resume(job_id)
        except (KeyError, OpenSWEError) as exc:
            raise HTTPException(status_code=409, detail=str(exc)) from exc
        return RedirectResponse(f"/jobs/{job_id}", status_code=303)

    @router.post("/api/jobs/{job_id}/dependencies")
    async def add_dependency(
        job_id: str,
        file: Annotated[UploadFile, File()],
    ) -> dict[str, Any]:
        dependency = await _write_dependency(file, store, job_id, settings.max_upload_mb)
        try:
            runner.follow_up(
                job_id,
                dependency=dependency,
                media_type=file.content_type or "application/octet-stream",
            )
            runner.resume(job_id)
        except (KeyError, OpenSWEError) as exc:
            raise HTTPException(status_code=409, detail=str(exc)) from exc
        return store.public_job(job_id) or {}

    @router.get("/api/jobs/{job_id}/artifacts")
    def list_artifacts(job_id: str) -> list[dict[str, str]]:
        job = _get_job_or_404(store, job_id)
        return [
            {
                "id": artifact.id,
                "name": artifact.name,
                "kind": artifact.kind,
                "media_type": artifact.media_type,
                "download": f"/jobs/{job_id}/artifacts/{artifact.id}",
            }
            for artifact in job.artifacts
        ]

    @router.get("/jobs/{job_id}/artifacts/{artifact_id}")
    def download_artifact(job_id: str, artifact_id: str) -> FileResponse:
        job = _get_job_or_404(store, job_id)
        artifact = next(
            (candidate for candidate in job.artifacts if candidate.id == artifact_id),
            None,
        )
        if artifact is None or not artifact.path.is_file():
            raise HTTPException(status_code=404, detail="Requested artifact is missing")
        return FileResponse(
            artifact.path,
            media_type=artifact.media_type,
            filename=artifact.name,
            headers={"X-Content-Type-Options": "nosniff"},
        )

    @router.get("/jobs/{job_id}/download")
    def download_ods(job_id: str) -> FileResponse:
        job = _completed_job_with_file(store, job_id, "ods")
        return FileResponse(
            job.output_path,
            media_type="application/vnd.oasis.opendocument.spreadsheet",
            filename=f"{safe_download_stem(job.original_filename)}.ods",
        )

    @router.get("/api/jobs/{job_id}/download")
    def download_ods_api(job_id: str) -> FileResponse:
        return download_ods(job_id)

    @router.get("/jobs/{job_id}/report.json")
    def download_report_json(job_id: str) -> FileResponse:
        job = _completed_job_with_file(store, job_id, "json")
        return FileResponse(
            job.report_json_path,
            media_type="application/json",
            filename=f"{safe_download_stem(job.original_filename)}-xlsliberator-report.json",
        )

    @router.get("/jobs/{job_id}/report.md")
    def download_report_md(job_id: str) -> FileResponse:
        job = _completed_job_with_file(store, job_id, "md")
        return FileResponse(
            job.report_md_path,
            media_type="text/markdown",
            filename=f"{safe_download_stem(job.original_filename)}-xlsliberator-report.md",
        )

    return router


async def _handle_upload(
    request: Request,
    background_tasks: BackgroundTasks,
    file: UploadFile,
    store: JobStore,
    runner: WebJobRunner,
    settings: WebSettings,
    *,
    force_json: bool = False,
) -> Response:
    try:
        original_filename = file.filename or ""
        extension = validate_upload_filename(original_filename)
        job_id = generate_job_id()
        paths = safe_job_paths(settings.data_dir, job_id, original_filename)
        paths.job_dir.mkdir(parents=True, exist_ok=False)
        bytes_written = await _write_upload(file, paths.input_path, settings.max_upload_mb)
        validate_upload_signature(paths.input_path, extension)
    except UploadValidationError as exc:
        if "paths" in locals():
            import shutil

            shutil.rmtree(paths.job_dir, ignore_errors=True)
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    store.create_job(
        job_id=job_id,
        original_filename=Path(original_filename).name,
        input_path=paths.input_path,
        output_path=paths.output_path,
        report_json_path=paths.report_json_path,
        report_md_path=paths.report_md_path,
        log_bundle_path=paths.log_bundle_path,
        profile_dir=paths.profile_dir,
    )
    store.add_event(
        job_id,
        phase=JobPhase.UPLOADED,
        step="uploaded",
        message="Upload received",
        percent=5,
        details={"bytes": bytes_written},
    )
    store.add_event(job_id, phase=JobPhase.QUEUED, step="queued", message="Job queued", percent=8)
    background_tasks.add_task(runner.submit, job_id)

    wants_json = force_json or "application/json" in request.headers.get("accept", "")
    if wants_json:
        return JSONResponse(store.public_job(job_id), status_code=202)
    return RedirectResponse(f"/jobs/{job_id}", status_code=303)


async def _write_upload(file: UploadFile, destination: Path, max_upload_mb: int) -> int:
    max_bytes = max_upload_mb * 1024 * 1024
    total = 0
    with destination.open("wb") as handle:
        while chunk := await file.read(1024 * 1024):
            total += len(chunk)
            if total > max_bytes:
                raise UploadValidationError(f"Upload exceeds {max_upload_mb} MB limit")
            handle.write(chunk)
    return total


async def _write_dependency(
    file: UploadFile,
    store: JobStore,
    job_id: str,
    max_upload_mb: int,
) -> Path:
    job = _get_job_or_404(store, job_id)
    filename = file.filename or ""
    if (
        not filename
        or Path(filename).name != filename
        or any(character in filename for character in ("/", "\\", "\x00"))
    ):
        raise HTTPException(status_code=400, detail="Invalid dependency filename")
    dependency_dir = job.input_path.parent / "dependencies"
    dependency_dir.mkdir(parents=True, exist_ok=True)
    destination = dependency_dir / filename
    try:
        await _write_upload(file, destination, max_upload_mb)
    except UploadValidationError as exc:
        destination.unlink(missing_ok=True)
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    return destination


def _get_job_or_404(store: JobStore, job_id: str) -> WebJob:
    job = store.get_job(job_id)
    if job is None:
        raise HTTPException(status_code=404, detail="Unknown job")
    return job


def _completed_job_with_file(store: JobStore, job_id: str, kind: str) -> WebJob:
    job = _get_job_or_404(store, job_id)
    if job.status != JobPhase.COMPLETED or job.operation_status != GateExecutionStatus.PASSED:
        raise HTTPException(status_code=409, detail="Job is not complete")
    path = {
        "ods": job.output_path,
        "json": job.report_json_path,
        "md": job.report_md_path,
    }[kind]
    if not path.exists():
        raise HTTPException(status_code=404, detail="Requested artifact is missing")
    return job


def _load_showcase_replay(job: WebJob) -> dict[str, Any]:
    recordings = [artifact for artifact in job.artifacts if artifact.kind == "showcase-recording"]
    results = [artifact for artifact in job.artifacts if artifact.kind == "showcase-result"]
    if len(recordings) != 1 or len(results) != 1:
        raise HTTPException(status_code=409, detail="Showcase evidence manifest is ambiguous")
    recording = recordings[0]
    result = results[0]
    if recording.media_type != "video/webm" or result.media_type != "application/json":
        raise HTTPException(status_code=409, detail="Showcase evidence media types are invalid")
    if not _is_job_artifact(job, recording.path) or not _is_job_artifact(job, result.path):
        raise HTTPException(status_code=409, detail="Showcase evidence path is invalid")
    if not recording.path.is_file() or recording.path.stat().st_size == 0:
        raise HTTPException(status_code=404, detail="Showcase recording is missing")
    if (
        not result.path.is_file()
        or result.path.stat().st_size == 0
        or result.path.stat().st_size > 2 * 1024 * 1024
    ):
        raise HTTPException(status_code=409, detail="Showcase result is missing or oversized")
    try:
        raw = json.loads(result.path.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise HTTPException(status_code=409, detail="Showcase result is not valid JSON") from exc
    if not isinstance(raw, dict) or raw.get("status") != "passed":
        raise HTTPException(status_code=409, detail="Showcase result did not pass")
    scenario_id = raw.get("scenario_id")
    if (
        not isinstance(scenario_id, str)
        or re.fullmatch(r"[A-Za-z0-9_.-]{1,100}", scenario_id) is None
    ):
        raise HTTPException(status_code=409, detail="Showcase scenario identifier is invalid")
    raw_operations = raw.get("operations")
    if not isinstance(raw_operations, list) or not 1 <= len(raw_operations) <= 100:
        raise HTTPException(status_code=409, detail="Showcase operation evidence is invalid")

    operations: list[dict[str, int | float | str]] = []
    for expected_sequence, raw_operation in enumerate(raw_operations, start=1):
        if not isinstance(raw_operation, dict):
            raise HTTPException(status_code=409, detail="Showcase operation is invalid")
        sequence = raw_operation.get("sequence")
        kind = raw_operation.get("kind")
        status = raw_operation.get("status")
        duration_ms = raw_operation.get("duration_ms")
        if (
            not isinstance(sequence, int)
            or isinstance(sequence, bool)
            or sequence != expected_sequence
        ):
            raise HTTPException(status_code=409, detail="Showcase operation sequence is invalid")
        if not isinstance(kind, str) or re.fullmatch(r"[a-z][a-z0-9_-]{0,39}", kind) is None:
            raise HTTPException(status_code=409, detail="Showcase operation kind is invalid")
        if status != "passed":
            raise HTTPException(status_code=409, detail="Showcase operation did not pass")
        if (
            not isinstance(duration_ms, (int, float))
            or isinstance(duration_ms, bool)
            or not math.isfinite(duration_ms)
            or not 0 <= duration_ms <= 300_000
        ):
            raise HTTPException(status_code=409, detail="Showcase operation duration is invalid")
        operations.append(
            {
                "sequence": sequence,
                "kind": kind,
                "status": status,
                "duration_ms": duration_ms,
            }
        )
    return {
        "scenario_id": scenario_id,
        "operations": operations,
        "recording_url": f"/jobs/{job.id}/artifacts/{recording.id}",
    }


def _is_job_artifact(job: WebJob, path: Path) -> bool:
    try:
        path.resolve().relative_to(job.input_path.parent.resolve())
    except (OSError, ValueError):
        return False
    return True


def _load_report_summary(job: WebJob) -> dict[str, Any] | None:
    if not job.report_json_path.exists():
        return None
    try:
        raw = json.loads(job.report_json_path.read_text())
    except (OSError, json.JSONDecodeError):
        return None
    warnings = raw.get("warnings", [])
    errors = raw.get("errors", [])
    return {
        "success": raw.get("success"),
        "duration_seconds": raw.get("duration_seconds", 0),
        "sheet_count": raw.get("sheet_count", 0),
        "total_cells": raw.get("total_cells", 0),
        "total_formulas": raw.get("total_formulas", 0),
        "formula_match_rate": raw.get("formula_match_rate", 0),
        "warnings_count": len(warnings),
        "errors_count": len(errors),
        "warnings": warnings[:5],
        "errors": errors[:5],
        "vba_modules": raw.get("vba_modules", 0),
        "vba_procedures": raw.get("vba_procedures", 0),
        "macro_functions_tested": raw.get("macro_functions_tested", 0),
        "macro_functions_failed": raw.get("macro_functions_failed", 0),
    }
