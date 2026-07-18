"""Thread-safe in-memory web job store."""

from __future__ import annotations

from datetime import UTC, datetime
from enum import StrEnum
from pathlib import Path
from threading import RLock
from typing import Any, Literal

from pydantic import BaseModel, Field

from xlsliberator.validation_models import GateExecutionStatus


class JobPhase(StrEnum):
    """Web conversion lifecycle phases."""

    UPLOADED = "uploaded"
    QUEUED = "queued"
    ANALYZING = "analyzing"
    CONVERTING = "converting"
    TRANSLATING = "translating"
    VERIFYING = "verifying"
    COMPLETED = "completed"
    FAILED = "failed"
    CANCELLED = "cancelled"


class JobEvent(BaseModel):
    """A single user-visible job event."""

    job_id: str
    phase: JobPhase
    step: str
    message: str
    percent: int | None = None
    level: Literal["info", "warning", "error"] = "info"
    timestamp: datetime = Field(default_factory=lambda: datetime.now(UTC))
    details: dict[str, Any] = Field(default_factory=dict)


class WebJob(BaseModel):
    """Internal web job state."""

    id: str
    original_filename: str
    input_path: Path
    output_path: Path
    report_json_path: Path
    report_md_path: Path
    log_bundle_path: Path
    profile_dir: Path
    status: JobPhase
    events: list[JobEvent] = Field(default_factory=list)
    created_at: datetime = Field(default_factory=lambda: datetime.now(UTC))
    updated_at: datetime = Field(default_factory=lambda: datetime.now(UTC))
    error: str | None = None
    cancellation_requested: bool = False
    operation_status: GateExecutionStatus = GateExecutionStatus.NOT_RUN


class JobStore:
    """In-memory job store protected by a lock."""

    def __init__(self) -> None:
        self._jobs: dict[str, WebJob] = {}
        self._lock = RLock()

    def create_job(
        self,
        *,
        job_id: str,
        original_filename: str,
        input_path: Path,
        output_path: Path,
        report_json_path: Path,
        report_md_path: Path,
        log_bundle_path: Path,
        profile_dir: Path,
    ) -> WebJob:
        """Create and store a new job."""
        with self._lock:
            job = WebJob(
                id=job_id,
                original_filename=original_filename,
                input_path=input_path,
                output_path=output_path,
                report_json_path=report_json_path,
                report_md_path=report_md_path,
                log_bundle_path=log_bundle_path,
                profile_dir=profile_dir,
                status=JobPhase.UPLOADED,
            )
            self._jobs[job_id] = job
            return job.model_copy(deep=True)

    def get_job(self, job_id: str) -> WebJob | None:
        """Return a defensive copy of a job."""
        with self._lock:
            job = self._jobs.get(job_id)
            return job.model_copy(deep=True) if job is not None else None

    def add_event(
        self,
        job_id: str,
        *,
        phase: JobPhase | str,
        step: str,
        message: str,
        percent: int | None = None,
        level: Literal["info", "warning", "error"] = "info",
        details: dict[str, Any] | None = None,
    ) -> JobEvent:
        """Append an event and update job status."""
        with self._lock:
            job = self._require_job(job_id)
            normalized = JobPhase(phase)
            event = JobEvent(
                job_id=job_id,
                phase=normalized,
                step=step,
                message=message,
                percent=percent,
                level=level,
                details=details or {},
            )
            job.events.append(event)
            job.status = normalized
            job.updated_at = event.timestamp
            return event.model_copy(deep=True)

    def mark_failed(self, job_id: str, error: str) -> WebJob:
        """Mark a job as failed."""
        self.add_event(job_id, phase=JobPhase.FAILED, step="failed", message=error, level="error")
        with self._lock:
            job = self._require_job(job_id)
            job.error = error
            job.operation_status = GateExecutionStatus.FAILED
            return job.model_copy(deep=True)

    def mark_completed(self, job_id: str) -> WebJob:
        """Mark a job as completed."""
        self.add_event(
            job_id,
            phase=JobPhase.COMPLETED,
            step="completed",
            message="Conversion complete",
            percent=100,
        )
        with self._lock:
            job = self._require_job(job_id)
            job.operation_status = GateExecutionStatus.PASSED
            return job.model_copy(deep=True)

    def request_cancel(self, job_id: str) -> WebJob:
        """Request best-effort cancellation."""
        with self._lock:
            job = self._require_job(job_id)
            if job.status in {JobPhase.COMPLETED, JobPhase.FAILED, JobPhase.CANCELLED}:
                raise ValueError(f"Cannot cancel job in {job.status.value} state")
            job.cancellation_requested = True
            if job.status == JobPhase.QUEUED:
                job.status = JobPhase.CANCELLED
                job.operation_status = GateExecutionStatus.SKIPPED
            job.updated_at = datetime.now(UTC)
        self.add_event(
            job_id,
            phase=JobPhase.CANCELLED if job.status == JobPhase.CANCELLED else job.status,
            step="cancel",
            message="Cancellation requested",
            level="warning",
        )
        return self.get_job(job_id) or job

    def list_jobs(self, limit: int = 50) -> list[WebJob]:
        """List recent jobs."""
        with self._lock:
            jobs = sorted(self._jobs.values(), key=lambda job: job.created_at, reverse=True)
            return [job.model_copy(deep=True) for job in jobs[:limit]]

    def public_job(self, job_id: str) -> dict[str, Any] | None:
        """Return JSON-safe public job data without internal paths."""
        job = self.get_job(job_id)
        if job is None:
            return None
        return public_job_dict(job)

    def _require_job(self, job_id: str) -> WebJob:
        job = self._jobs.get(job_id)
        if job is None:
            raise KeyError(job_id)
        return job


def public_job_dict(job: WebJob) -> dict[str, Any]:
    """Serialize a job without leaking server filesystem paths."""
    return {
        "id": job.id,
        "original_filename": job.original_filename,
        "status": job.status.value,
        "transport_success": True,
        "operation_status": job.operation_status.value,
        "created_at": job.created_at.isoformat(),
        "updated_at": job.updated_at.isoformat(),
        "error": job.error,
        "cancellation_requested": job.cancellation_requested,
        "events": [
            {
                "index": index,
                "job_id": event.job_id,
                "phase": event.phase.value,
                "step": event.step,
                "message": event.message,
                "percent": event.percent,
                "level": event.level,
                "timestamp": event.timestamp.isoformat(),
                "details": _public_details(event.details),
            }
            for index, event in enumerate(job.events)
        ],
        "downloads": _download_links(job),
    }


def _download_links(job: WebJob) -> dict[str, str]:
    if job.status != JobPhase.COMPLETED or job.operation_status != GateExecutionStatus.PASSED:
        return {}
    return {
        "ods": f"/jobs/{job.id}/download",
        "report_json": f"/jobs/{job.id}/report.json",
        "report_md": f"/jobs/{job.id}/report.md",
    }


def _public_details(details: dict[str, Any]) -> dict[str, Any]:
    public: dict[str, Any] = {}
    for key, value in details.items():
        lowered = key.lower()
        if lowered in {"input", "output", "file"} or "path" in lowered or isinstance(value, Path):
            continue
        public[key] = value
    return public
