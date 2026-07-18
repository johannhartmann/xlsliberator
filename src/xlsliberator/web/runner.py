"""Background conversion runner for web jobs."""

from __future__ import annotations

import shutil
import zipfile
from concurrent.futures import Future, ThreadPoolExecutor
from pathlib import Path
from typing import Any

from loguru import logger

from xlsliberator.api import convert
from xlsliberator.report import ConversionReport
from xlsliberator.web.jobs import JobPhase, JobStore
from xlsliberator.web.schemas import WebSettings
from xlsliberator.web.security import safe_download_stem


class WebJobRunner:
    """Runs conversion jobs with bounded concurrency."""

    def __init__(self, store: JobStore, settings: WebSettings) -> None:
        self.store = store
        self.settings = settings
        self._executor = ThreadPoolExecutor(max_workers=settings.worker_count)

    def submit(self, job_id: str) -> Future[None]:
        """Submit a job to the executor."""
        return self._executor.submit(self.run_job, job_id)

    def run_job(self, job_id: str) -> None:
        """Run one conversion job and record status events."""
        job = self.store.get_job(job_id)
        if job is None:
            logger.warning(f"Unknown web conversion job: {job_id}")
            return
        if job.cancellation_requested or job.status == JobPhase.CANCELLED:
            self.store.add_event(
                job_id,
                phase=JobPhase.CANCELLED,
                step="cancelled",
                message="Job cancelled before conversion started",
                level="warning",
            )
            return

        job.profile_dir.mkdir(parents=True, exist_ok=True)
        try:
            self.store.add_event(
                job_id,
                phase=JobPhase.ANALYZING,
                step="analyzing",
                message="Analyzing workbook",
                percent=10,
            )

            def progress(phase: str, message: str, details: dict[str, Any]) -> None:
                mapped = _phase_from_core(phase)
                self.store.add_event(
                    job_id,
                    phase=mapped,
                    step=phase,
                    message=message,
                    details=details,
                )

            report = convert(
                job.input_path,
                job.output_path,
                strict=False,
                embed_macros=self.settings.embed_macros,
                use_agent=self.settings.use_agent,
                validate_macro_execution=False,
                allow_global_macro_security_change=False,
                progress_callback=progress,
                user_installation_dir=job.profile_dir,
            )
            report.input_file = job.original_filename
            report.output_file = f"{safe_download_stem(job.original_filename)}.ods"
            _write_reports(report, job.report_json_path, job.report_md_path)
            valid_output = job.output_path.is_file() and zipfile.is_zipfile(job.output_path)
            if report.success and valid_output:
                self.store.mark_completed(job_id)
            else:
                if not valid_output and not report.errors:
                    report.errors.append("Conversion output is not a valid ODS ZIP package")
                    _write_reports(report, job.report_json_path, job.report_md_path)
                error = report.errors[-1] if report.errors else "Conversion failed"
                self.store.mark_failed(job_id, error)
        except Exception as exc:
            logger.exception(f"Web conversion job failed: {job_id}")
            self.store.mark_failed(job_id, str(exc))


def _write_reports(report: ConversionReport, json_path: Path, markdown_path: Path) -> None:
    json_path.parent.mkdir(parents=True, exist_ok=True)
    report.save_json(json_path)
    report.save_markdown(markdown_path)


def _phase_from_core(phase: str) -> JobPhase:
    if phase in {"converting", "repairing"}:
        return JobPhase.CONVERTING
    if phase in {"extracting_vba", "translating", "embedding"}:
        return JobPhase.TRANSLATING
    if phase.startswith("verifying"):
        return JobPhase.VERIFYING
    if phase == "completed":
        # A core progress event is not terminal proof. The runner promotes the
        # job only after the report and output package have both been checked.
        return JobPhase.VERIFYING
    if phase == "failed":
        return JobPhase.FAILED
    return JobPhase.ANALYZING


def cleanup_profile(profile_dir: Path) -> None:
    """Remove a per-job LibreOffice profile directory if present."""
    if profile_dir.exists():
        shutil.rmtree(profile_dir)
