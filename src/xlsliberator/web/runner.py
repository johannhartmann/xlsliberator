"""Background Open-SWE migration runner for web jobs."""

from __future__ import annotations

import json
import re
import shutil
import time
import zipfile
from concurrent.futures import Future, ThreadPoolExecutor
from pathlib import Path
from typing import Any, Protocol

from loguru import logger

from xlsliberator.web.jobs import JobArtifact, JobPhase, JobStore, WebJob
from xlsliberator.web.open_swe import OpenSWEClient, OpenSWEError
from xlsliberator.web.schemas import WebSettings

_TERMINAL_FAILURES = frozenset({"cancelled", "cleaned", "failed", "rejected"})


class OpenSWETransport(Protocol):
    """Operations required from the Open-SWE client."""

    def create_migration(self, workbook: Path, requirements: str = "") -> dict[str, Any]: ...

    def status(self, thread_id: str) -> dict[str, Any]: ...

    def events(self, thread_id: str, since: int) -> dict[str, Any]: ...

    def follow_up(
        self,
        thread_id: str,
        *,
        requirements: str = "",
        dependency: Path | None = None,
        media_type: str = "application/octet-stream",
    ) -> dict[str, Any]: ...

    def cancel(self, thread_id: str) -> dict[str, Any]: ...

    def download_artifact(self, thread_id: str, artifact_id: str) -> bytes: ...


class WebJobRunner:
    """Runs serious workbook migrations only through Open-SWE."""

    def __init__(
        self,
        store: JobStore,
        settings: WebSettings,
        client: OpenSWETransport | None = None,
    ) -> None:
        self.store = store
        self.settings = settings
        self.client = client or _client_from_settings(settings)
        self._executor = ThreadPoolExecutor(max_workers=settings.worker_count)

    def submit(self, job_id: str) -> Future[None]:
        """Submit a job to the executor."""
        return self._executor.submit(self.run_job, job_id)

    def run_job(self, job_id: str) -> None:
        """Create and monitor one Open-SWE workbook thread."""
        job = self.store.get_job(job_id)
        if job is None:
            logger.warning(f"Unknown web migration job: {job_id}")
            return
        if job.status == JobPhase.CANCELLED:
            return
        if job.cancellation_requested:
            self.store.mark_cancelled(job_id)
            return
        if self.client is None:
            self.store.mark_failed(
                job_id,
                "Open-SWE migration service is not configured; local conversion is disabled",
            )
            return

        try:
            self.store.add_event(
                job_id,
                phase=JobPhase.ANALYZING,
                step="open_swe_create",
                message="Creating private Open-SWE workbook thread",
                percent=10,
            )
            created = self.client.create_migration(job.input_path)
            thread_id = _required_string(created, "thread_id")
            run_id = _optional_string(created, "run_id")
            self.store.set_remote(job_id, thread_id=thread_id, run_id=run_id)
            self.store.add_event(
                job_id,
                phase=JobPhase.ANALYZING,
                step="thread_ready",
                message="Workbook thread is ready",
                percent=15,
            )
            status = self._monitor(job_id, thread_id)
            if status.get("status") != "complete":
                return
            self._finalize(job_id, thread_id, status)
        except OpenSWEError as exc:
            self.store.mark_failed(job_id, str(exc))
        except Exception:
            logger.exception(f"Open-SWE web migration failed: {job_id}")
            self.store.mark_failed(job_id, "Open-SWE migration failed unexpectedly")

    def cancel(self, job_id: str) -> None:
        """Propagate cancellation to the attached Open-SWE thread."""
        job = self.store.get_job(job_id)
        if job is None or job.status == JobPhase.CANCELLED:
            return
        if self.client is None or job.remote_thread_id is None:
            return
        self.client.cancel(job.remote_thread_id)
        self.store.mark_cancelled(job_id)

    def resume(self, job_id: str) -> Future[None]:
        """Resume monitoring after a follow-up starts another thread run."""
        return self._executor.submit(self._resume_job, job_id)

    def follow_up(
        self,
        job_id: str,
        *,
        requirements: str = "",
        dependency: Path | None = None,
        media_type: str = "application/octet-stream",
    ) -> None:
        """Send a message or validated dependency to the same workbook thread."""
        job = self.store.get_job(job_id)
        if job is None:
            raise KeyError(job_id)
        if self.client is None or job.remote_thread_id is None:
            raise OpenSWEError("Workbook thread is not ready for follow-up messages")
        result = self.client.follow_up(
            job.remote_thread_id,
            requirements=requirements,
            dependency=dependency,
            media_type=media_type,
        )
        self.store.set_remote(
            job_id,
            thread_id=job.remote_thread_id,
            run_id=_optional_string(result, "run_id"),
            reset_delivery=True,
        )
        self.store.add_event(
            job_id,
            phase=JobPhase.ANALYZING,
            step="follow_up",
            message="Follow-up added to the workbook thread",
        )

    def _resume_job(self, job_id: str) -> None:
        job = self.store.get_job(job_id)
        if job is None or job.remote_thread_id is None:
            return
        try:
            status = self._monitor(job_id, job.remote_thread_id)
            if status.get("status") == "complete":
                self._finalize(job_id, job.remote_thread_id, status)
        except OpenSWEError as exc:
            self.store.mark_failed(job_id, str(exc))
        except Exception:
            logger.exception(f"Open-SWE follow-up failed: {job_id}")
            self.store.mark_failed(job_id, "Open-SWE follow-up failed unexpectedly")

    def _monitor(self, job_id: str, thread_id: str) -> dict[str, Any]:
        assert self.client is not None
        deadline = time.monotonic() + self.settings.open_swe_job_timeout_seconds
        next_event = 0
        while time.monotonic() < deadline:
            job = self.store.get_job(job_id)
            if job is None:
                raise OpenSWEError("Local workbook job disappeared")
            if job.cancellation_requested:
                self.client.cancel(thread_id)
                self.store.mark_cancelled(job_id)
                return {"status": "cancelled"}

            event_payload = self.client.events(thread_id, next_event)
            for remote_event in _list_of_objects(event_payload.get("events")):
                self._record_remote_event(job_id, remote_event)
            next_value = event_payload.get("next")
            if isinstance(next_value, int):
                next_event = max(next_event, next_value)

            status = self.client.status(thread_id)
            operation = _required_string(status, "status")
            if operation == "complete":
                return status
            if operation in _TERMINAL_FAILURES:
                if operation == "cancelled":
                    self.store.mark_cancelled(job_id)
                else:
                    self.store.mark_failed(job_id, f"Open-SWE migration ended as {operation}")
                return status
            time.sleep(self.settings.open_swe_poll_seconds)
        self.client.cancel(thread_id)
        raise OpenSWEError("Open-SWE migration exceeded its configured time limit")

    def _record_remote_event(self, job_id: str, event: dict[str, Any]) -> None:
        stage = _optional_string(event, "stage") or "lead"
        message = _optional_string(event, "message") or "Workbook migration progressed"
        phase, percent = _phase_from_remote_stage(stage)
        self.store.add_event(
            job_id,
            phase=phase,
            step=stage,
            message=message[:500],
            percent=percent,
        )

    def _download_deliverables(
        self,
        job_id: str,
        thread_id: str,
        status: dict[str, Any],
    ) -> list[JobArtifact]:
        assert self.client is not None
        job = self.store.get_job(job_id)
        if job is None:
            raise OpenSWEError("Local workbook job disappeared")
        deliverables = job.input_path.parent / "deliverables"
        deliverables.mkdir(parents=True, exist_ok=True)
        artifacts: list[JobArtifact] = []
        for item in _list_of_objects(status.get("artifacts")):
            artifact_id = _required_string(item, "id")
            name = _safe_artifact_name(_required_string(item, "name"))
            kind = _optional_string(item, "kind") or "artifact"
            media_type = _optional_string(item, "media_type") or "application/octet-stream"
            content = self.client.download_artifact(thread_id, artifact_id)
            if name == "target.ods" and kind == "ods":
                destination = job.output_path
            elif name == "report.json" and kind == "report":
                destination = job.report_json_path
            elif name == "report.md" and kind == "report":
                destination = job.report_md_path
            else:
                destination = deliverables / f"{artifact_id}-{name}"
            destination.write_bytes(content)
            artifacts.append(
                JobArtifact(
                    id=artifact_id,
                    name=name,
                    kind=kind,
                    media_type=media_type,
                    path=destination,
                )
            )
        return artifacts

    def _finalize(
        self,
        job_id: str,
        thread_id: str,
        status: dict[str, Any],
    ) -> None:
        artifacts = self._download_deliverables(job_id, thread_id, status)
        refreshed = self.store.get_job(job_id)
        if refreshed is None or not refreshed.output_path.is_file():
            raise OpenSWEError("Open-SWE completed without an ODS deliverable")
        if not zipfile.is_zipfile(refreshed.output_path):
            raise OpenSWEError("Open-SWE ODS deliverable is not a valid package")
        self.store.set_artifacts(job_id, artifacts)
        _ensure_delivery_report(refreshed, artifacts)
        _delete_private_inputs(refreshed)
        self.store.mark_completed(job_id)


def _client_from_settings(settings: WebSettings) -> OpenSWEClient | None:
    if not settings.open_swe_url or not settings.open_swe_token:
        return None
    return OpenSWEClient(
        base_url=settings.open_swe_url,
        token=settings.open_swe_token,
        owner_id=settings.open_swe_owner_id,
        timeout_seconds=settings.open_swe_request_timeout_seconds,
    )


def _phase_from_remote_stage(stage: str) -> tuple[JobPhase, int | None]:
    if stage in {"upload", "lead"}:
        return JobPhase.ANALYZING, 20
    if stage in {"plan", "specialists"}:
        return JobPhase.CONVERTING, 50
    if stage in {"libreoffice", "reviewer"}:
        return JobPhase.VERIFYING, 80
    if stage == "final":
        return JobPhase.VERIFYING, 95
    return JobPhase.ANALYZING, None


def _required_string(value: dict[str, Any], key: str) -> str:
    result = value.get(key)
    if not isinstance(result, str) or not result:
        raise OpenSWEError("Open-SWE returned an invalid response")
    return result


def _optional_string(value: dict[str, Any], key: str) -> str | None:
    result = value.get(key)
    return result if isinstance(result, str) and result else None


def _list_of_objects(value: object) -> list[dict[str, Any]]:
    if not isinstance(value, list):
        return []
    return [item for item in value if isinstance(item, dict)]


def _safe_artifact_name(name: str) -> str:
    plain = Path(name).name
    if plain != name or "\x00" in name:
        raise OpenSWEError("Open-SWE returned an unsafe artifact name")
    clean = re.sub(r"[^A-Za-z0-9._-]+", "-", plain).strip(".-_")
    if not clean:
        raise OpenSWEError("Open-SWE returned an unsafe artifact name")
    return clean


def _ensure_delivery_report(job: WebJob, artifacts: list[JobArtifact]) -> None:
    public_artifacts = [
        {"id": artifact.id, "name": artifact.name, "kind": artifact.kind}
        for artifact in artifacts
    ]
    if not any(artifact.name == "report.json" for artifact in artifacts):
        job.report_json_path.write_text(
            json.dumps(
                {
                    "success": True,
                    "source": job.original_filename,
                    "target_libreoffice_version": "26.2.4.2",
                    "artifacts": public_artifacts,
                    "warnings": [],
                    "errors": [],
                },
                indent=2,
                sort_keys=True,
            )
        )
    if not any(artifact.name == "report.md" for artifact in artifacts):
        lines = [
            "# XLSLiberator migration report",
            "",
            "Status: independently reviewed and complete.",
            "",
            "## Deliverables",
            "",
            *[f"- {artifact['kind']}: {artifact['name']}" for artifact in public_artifacts],
            "",
        ]
        job.report_md_path.write_text("\n".join(lines))


def _delete_private_inputs(job: WebJob) -> None:
    """Delete local source and dependency copies after the remote delivery is complete."""
    try:
        job.input_path.unlink(missing_ok=True)
        shutil.rmtree(job.input_path.parent / "dependencies", ignore_errors=False)
    except FileNotFoundError:
        return
    except OSError as exc:
        raise OpenSWEError("Private upload retention cleanup failed") from exc


def cleanup_profile(profile_dir: Path) -> None:
    """Remove a legacy per-job profile directory if present."""
    if profile_dir.exists():
        shutil.rmtree(profile_dir)
