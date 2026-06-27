"""FastAPI application factory."""

from __future__ import annotations

import shutil
import subprocess
from pathlib import Path
from typing import Any

from fastapi import FastAPI
from fastapi.staticfiles import StaticFiles

from xlsliberator.web.cleanup import cleanup_old_jobs
from xlsliberator.web.jobs import JobStore
from xlsliberator.web.routes import create_router
from xlsliberator.web.runner import WebJobRunner
from xlsliberator.web.schemas import WebSettings


def create_app(settings: WebSettings | None = None) -> FastAPI:
    """Create the XLSLiberator web application."""
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
    """Return readiness checks without raising when optional tools are absent."""
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    version = None
    if soffice:
        try:
            result = subprocess.run(
                [soffice, "--version"],
                capture_output=True,
                text=True,
                timeout=5,
                check=False,
            )
            version = (result.stdout or result.stderr).strip() or None
        except (OSError, subprocess.SubprocessError):
            version = None
    return {
        "data_dir_writable": _is_writable(settings.data_dir),
        "soffice_available": bool(soffice),
        "version": version,
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
