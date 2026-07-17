"""FastAPI application factory."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from fastapi import FastAPI
from fastapi.staticfiles import StaticFiles

from xlsliberator.container_boundary import require_application_container
from xlsliberator.docker_runtime import DockerRuntimeUnavailable, LibreOfficeDockerRuntime
from xlsliberator.web.cleanup import cleanup_old_jobs
from xlsliberator.web.jobs import JobStore
from xlsliberator.web.routes import create_router
from xlsliberator.web.runner import WebJobRunner
from xlsliberator.web.schemas import WebSettings


def create_app(settings: WebSettings | None = None) -> FastAPI:
    """Create the XLSLiberator web application."""
    require_application_container()
    resolved_settings = settings or WebSettings.from_env()
    resolved_settings.data_dir.mkdir(parents=True, exist_ok=True)
    store = JobStore()
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
    def startup_cleanup() -> None:
        from datetime import timedelta

        cleanup_old_jobs(
            resolved_settings.data_dir,
            timedelta(hours=resolved_settings.job_retention_hours),
        )

    return app


def readiness(settings: WebSettings) -> dict[str, Any]:
    """Return Docker runtime readiness without inspecting host office software."""
    runtime_available = False
    image_id = None
    version = None
    error = None
    try:
        identity = LibreOfficeDockerRuntime().resolve_identity()
        runtime_available = True
        image_id = identity.image_id
        version = identity.version
    except DockerRuntimeUnavailable as exc:
        error = str(exc)
    return {
        "data_dir_writable": _is_writable(settings.data_dir),
        "docker_runtime_available": runtime_available,
        "image_id": image_id,
        "version": version,
        "runtime_error": error,
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
