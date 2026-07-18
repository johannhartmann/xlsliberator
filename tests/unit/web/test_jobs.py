from pathlib import Path

import pytest

from xlsliberator.validation_models import GateExecutionStatus
from xlsliberator.web.jobs import JobPhase, JobStore


def _create_job(store: JobStore, tmp_path: Path) -> str:
    job_id = "11111111-1111-4111-8111-111111111111"
    store.create_job(
        job_id=job_id,
        original_filename="book.xlsx",
        input_path=tmp_path / "input.xlsx",
        output_path=tmp_path / "output.ods",
        report_json_path=tmp_path / "report.json",
        report_md_path=tmp_path / "report.md",
        log_bundle_path=tmp_path / "logs.zip",
        profile_dir=tmp_path / "profile",
    )
    return job_id


def test_job_store_events_and_public_serialization(tmp_path: Path) -> None:
    store = JobStore()
    job_id = _create_job(store, tmp_path)

    store.add_event(job_id, phase=JobPhase.UPLOADED, step="upload", message="ok")
    store.add_event(job_id, phase=JobPhase.QUEUED, step="queue", message="queued")
    public = store.public_job(job_id)

    assert public is not None
    assert public["status"] == "queued"
    assert public["transport_success"] is True
    assert public["operation_status"] == GateExecutionStatus.NOT_RUN.value
    assert [event["index"] for event in public["events"]] == [0, 1]
    assert str(tmp_path) not in str(public)


def test_job_store_completion_and_download_links(tmp_path: Path) -> None:
    store = JobStore()
    job_id = _create_job(store, tmp_path)

    store.mark_completed(job_id)
    public = store.public_job(job_id)

    assert public is not None
    assert public["status"] == "completed"
    assert public["operation_status"] == GateExecutionStatus.PASSED.value
    assert public["downloads"]["ods"] == f"/jobs/{job_id}/download"


def test_progress_event_cannot_expose_downloads_without_operation_pass(tmp_path: Path) -> None:
    store = JobStore()
    job_id = _create_job(store, tmp_path)

    store.add_event(
        job_id,
        phase=JobPhase.COMPLETED,
        step="completed",
        message="Unverified progress event",
    )
    public = store.public_job(job_id)

    assert public is not None
    assert public["status"] == "completed"
    assert public["operation_status"] == GateExecutionStatus.NOT_RUN.value
    assert public["downloads"] == {}


def test_job_store_cancel_rules(tmp_path: Path) -> None:
    store = JobStore()
    job_id = _create_job(store, tmp_path)
    store.add_event(job_id, phase=JobPhase.QUEUED, step="queue", message="queued")

    cancelled = store.request_cancel(job_id)

    assert cancelled.status == JobPhase.CANCELLED
    assert cancelled.operation_status is GateExecutionStatus.SKIPPED
    with pytest.raises(ValueError):
        store.request_cancel(job_id)


def test_running_job_records_cancellation_request(tmp_path: Path) -> None:
    store = JobStore()
    job_id = _create_job(store, tmp_path)
    store.add_event(job_id, phase=JobPhase.CONVERTING, step="convert", message="running")

    job = store.request_cancel(job_id)

    assert job.cancellation_requested is True
    assert job.status == JobPhase.CONVERTING
    assert job.events[-1].step == "cancel"


def test_completed_job_cannot_cancel(tmp_path: Path) -> None:
    store = JobStore()
    job_id = _create_job(store, tmp_path)
    store.mark_completed(job_id)

    with pytest.raises(ValueError):
        store.request_cancel(job_id)
