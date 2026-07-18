"""FastAPI application factory."""

from __future__ import annotations

from datetime import timedelta
from pathlib import Path
from typing import Any

from fastapi import FastAPI
from fastapi.staticfiles import StaticFiles

from xlsliberator.container_boundary import require_application_container
from xlsliberator.web.cleanup import cleanup_old_jobs
from xlsliberator.web.jobs import JobPhase, JobStore
from xlsliberator.web.open_swe import OpenSWEClient
from xlsliberator.web.routes import create_router
from xlsliberator.web.runner import WebJobRunner
from xlsliberator.web.schemas import WebSettings


def create_app(settings: WebSettings | None = None) -> FastAPI:
    """Create the XLSLiberator web application."""
    require_application_container()
    resolved_settings = settings or WebSettings.from_env()
    resolved_settings.data_dir.mkdir(parents=True, exist_ok=True)
    cleanup_old_jobs(
        resolved_settings.data_dir,
        timedelta(hours=resolved_settings.job_retention_hours),
    )
    store = JobStore(resolved_settings.data_dir)
    runner = WebJobRunner(store, resolved_settings)

    app = FastAPI(title="XLSLiberator Web", version="0.1.0")
    app.state.settings = resolved_settings
    app.state.job_store = store
    app.state.job_runner = runner

    static_dir = Path(__file__).parent / "static"
    app.mount("/static", StaticFiles(directory=static_dir), name="static")
    app.include_router(create_router(store, runner, resolved_settings))

    @app.get("/healthz")
    def healthz() -> dict[str, str]:
        return {"status": "ok"}

    @app.get("/readyz")
    def readyz() -> dict[str, Any]:
        return readiness(resolved_settings)

    @app.on_event("startup")
    def startup_resume() -> None:
        for job in store.list_jobs(limit=10_000):
            if job.status in {
                JobPhase.COMPLETED,
                JobPhase.FAILED,
                JobPhase.CANCELLED,
            }:
                continue
            if job.remote_thread_id is None:
                runner.submit(job.id)
            else:
                runner.resume(job.id)

    return app


def readiness(settings: WebSettings) -> dict[str, Any]:
    """Return web workspace and Open-SWE service readiness."""
    configured = bool(settings.open_swe_url and settings.open_swe_token)
    reachable = False
    if configured:
        reachable = OpenSWEClient(
            base_url=settings.open_swe_url,
            token=settings.open_swe_token,
            owner_id=settings.open_swe_owner_id,
            timeout_seconds=min(2.0, settings.open_swe_request_timeout_seconds),
        ).ready()
    return {
        "data_dir_writable": _is_writable(settings.data_dir),
        "open_swe_configured": configured,
        "open_swe_reachable": reachable,
        "target_libreoffice_version": "26.2.4.2",
    }


def _is_writable(path: Path) -> bool:
    try:
        path.mkdir(parents=True, exist_ok=True)
        probe = path / ".write-test"
        probe.write_text("ok")
        probe.unlink()
        return True
    except OSError:
        return False
