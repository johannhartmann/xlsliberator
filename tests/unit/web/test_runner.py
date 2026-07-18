import io
import zipfile
from pathlib import Path
from typing import Any

from xlsliberator.validation_models import GateExecutionStatus
from xlsliberator.web.jobs import JobPhase, JobStore
from xlsliberator.web.runner import WebJobRunner, _phase_from_remote_stage
from xlsliberator.web.schemas import WebSettings


def _store_with_job(tmp_path: Path) -> tuple[JobStore, str]:
    store = JobStore()
    job_id = "22222222-2222-4222-8222-222222222222"
    input_path = tmp_path / "input.xlsx"
    input_path.write_bytes(b"PK\x03\x04content")
    store.create_job(
        job_id=job_id,
        original_filename="book.xlsx",
        input_path=input_path,
        output_path=tmp_path / "output.ods",
        report_json_path=tmp_path / "report.json",
        report_md_path=tmp_path / "report.md",
        log_bundle_path=tmp_path / "logs.zip",
        profile_dir=tmp_path / "profile",
    )
    store.add_event(job_id, phase=JobPhase.QUEUED, step="queued", message="queued")
    return store, job_id


def _ods_bytes() -> bytes:
    output = io.BytesIO()
    with zipfile.ZipFile(output, "w") as archive:
        archive.writestr("mimetype", "application/vnd.oasis.opendocument.spreadsheet")
    return output.getvalue()


class FakeOpenSWE:
    def __init__(self, status: str = "complete") -> None:
        self.operation_status = status
        self.cancelled: list[str] = []
        self.follow_ups: list[dict[str, Any]] = []

    def create_migration(self, workbook: Path, requirements: str = "") -> dict[str, Any]:
        assert workbook.name == "input.xlsx"
        assert requirements == ""
        return {"thread_id": "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa", "run_id": "run-1"}

    def status(self, thread_id: str) -> dict[str, Any]:
        assert thread_id
        return {
            "thread_id": thread_id,
            "status": self.operation_status,
            "artifacts": [
                {
                    "id": "a" * 24,
                    "name": "target.ods",
                    "kind": "ods",
                    "media_type": "application/vnd.oasis.opendocument.spreadsheet",
                },
                {
                    "id": "b" * 24,
                    "name": "bridge.py",
                    "kind": "generated",
                    "media_type": "text/x-python",
                },
                {
                    "id": "c" * 24,
                    "name": "save-reopen.json",
                    "kind": "evidence",
                    "media_type": "application/json",
                },
            ],
        }

    def events(self, thread_id: str, since: int) -> dict[str, Any]:
        assert thread_id
        events = [
            {
                "index": 0,
                "stage": "plan",
                "message": "Behavioral migration plan is ready",
                "status": "complete",
            },
            {
                "index": 1,
                "stage": "reviewer",
                "message": "Independent behavior review has reported",
                "status": "complete",
            },
        ]
        return {"events": events[since:], "next": len(events)}

    def follow_up(
        self,
        thread_id: str,
        *,
        requirements: str = "",
        dependency: Path | None = None,
        media_type: str = "application/octet-stream",
    ) -> dict[str, Any]:
        self.follow_ups.append(
            {
                "thread_id": thread_id,
                "requirements": requirements,
                "dependency": dependency,
                "media_type": media_type,
            }
        )
        return {"thread_id": thread_id, "run_id": "run-2"}

    def cancel(self, thread_id: str) -> dict[str, Any]:
        self.cancelled.append(thread_id)
        return {"thread_id": thread_id, "status": "cancelled"}

    def download_artifact(self, thread_id: str, artifact_id: str) -> bytes:
        assert thread_id
        return {
            "a" * 24: _ods_bytes(),
            "b" * 24: b"def migrate(): return True\n",
            "c" * 24: b'{"status":"passed"}',
        }[artifact_id]


def test_runner_uses_open_swe_and_downloads_delivery_bundle(tmp_path: Path) -> None:
    store, job_id = _store_with_job(tmp_path)
    fake = FakeOpenSWE()
    settings = WebSettings(
        data_dir=tmp_path,
        open_swe_poll_seconds=0.1,
        open_swe_job_timeout_seconds=30,
    )

    WebJobRunner(store, settings, client=fake).run_job(job_id)
    job = store.get_job(job_id)

    assert job is not None
    assert job.status == JobPhase.COMPLETED
    assert job.remote_thread_id == "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa"
    assert zipfile.is_zipfile(job.output_path)
    assert not job.input_path.exists()
    assert job.report_json_path.exists()
    assert {artifact.name for artifact in job.artifacts} == {
        "target.ods",
        "bridge.py",
        "save-reopen.json",
    }
    assert any(event.step == "reviewer" for event in job.events)


def test_runner_never_falls_back_to_local_conversion(tmp_path: Path) -> None:
    store, job_id = _store_with_job(tmp_path)

    WebJobRunner(store, WebSettings(data_dir=tmp_path)).run_job(job_id)
    job = store.get_job(job_id)

    assert job is not None
    assert job.status == JobPhase.FAILED
    assert job.error is not None
    assert "local conversion is disabled" in job.error


def test_runner_propagates_remote_failure(tmp_path: Path) -> None:
    store, job_id = _store_with_job(tmp_path)
    fake = FakeOpenSWE(status="failed")

    WebJobRunner(store, WebSettings(data_dir=tmp_path), client=fake).run_job(job_id)
    job = store.get_job(job_id)

    assert job is not None
    assert job.status == JobPhase.FAILED
    assert job.error == "Open-SWE migration ended as failed"


def test_follow_up_stays_on_attached_thread(tmp_path: Path) -> None:
    store, job_id = _store_with_job(tmp_path)
    fake = FakeOpenSWE()
    runner = WebJobRunner(store, WebSettings(data_dir=tmp_path), client=fake)
    store.set_remote(
        job_id,
        thread_id="aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa",
        run_id="run-1",
    )
    store.mark_completed(job_id)

    runner.follow_up(job_id, requirements="Preserve quarterly import behavior")

    assert fake.follow_ups[0]["thread_id"] == "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa"
    assert fake.follow_ups[0]["requirements"] == "Preserve quarterly import behavior"
    job = store.get_job(job_id)
    assert job is not None
    assert job.remote_run_id == "run-2"
    assert job.operation_status == GateExecutionStatus.NOT_RUN


def test_cancel_propagates_to_attached_thread(tmp_path: Path) -> None:
    store, job_id = _store_with_job(tmp_path)
    fake = FakeOpenSWE()
    runner = WebJobRunner(store, WebSettings(data_dir=tmp_path), client=fake)
    store.set_remote(
        job_id,
        thread_id="aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa",
        run_id="run-1",
    )

    runner.cancel(job_id)

    assert fake.cancelled == ["aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa"]
    job = store.get_job(job_id)
    assert job is not None
    assert job.status == JobPhase.CANCELLED


def test_remote_stages_map_to_non_terminal_local_phases() -> None:
    assert _phase_from_remote_stage("plan") == (JobPhase.CONVERTING, 50)
    assert _phase_from_remote_stage("reviewer") == (JobPhase.VERIFYING, 80)
    assert _phase_from_remote_stage("final") == (JobPhase.VERIFYING, 95)
