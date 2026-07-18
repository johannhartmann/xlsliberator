"""HTML and JSON routes for conversion jobs."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Annotated, Any

from fastapi import APIRouter, BackgroundTasks, File, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, RedirectResponse, Response
from fastapi.templating import Jinja2Templates

from xlsliberator.validation_models import GateExecutionStatus
from xlsliberator.web.jobs import JobPhase, JobStore, WebJob, public_job_dict
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
        return public_job_dict(job)

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


def _get_job_or_404(store: JobStore, job_id: str) -> WebJob:
    job = store.get_job(job_id)
    if job is None:
        raise HTTPException(status_code=404, detail="Unknown job")
    return job


def _completed_job_with_file(store: JobStore, job_id: str, kind: str) -> WebJob:
    job = _get_job_or_404(store, job_id)
    if (
        job.status != JobPhase.COMPLETED
        or job.operation_status != GateExecutionStatus.PASSED
    ):
        raise HTTPException(status_code=409, detail="Job is not complete")
    path = {
        "ods": job.output_path,
        "json": job.report_json_path,
        "md": job.report_md_path,
    }[kind]
    if not path.exists():
        raise HTTPException(status_code=404, detail="Requested artifact is missing")
    return job


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
