from pathlib import Path
from typing import Any

from xlsliberator.report import ConversionReport
from xlsliberator.web.jobs import JobPhase, JobStore
from xlsliberator.web.runner import WebJobRunner
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


def test_runner_calls_convert_and_writes_reports(tmp_path: Path, monkeypatch: Any) -> None:
    store, job_id = _store_with_job(tmp_path)
    calls: dict[str, Any] = {}

    def fake_convert(input_path: Path, output_path: Path, **kwargs: Any) -> ConversionReport:
        calls.update(kwargs)
        output_path.write_text("ods")
        kwargs["progress_callback"]("converting", "Converting", {})
        return ConversionReport(
            input_file=str(input_path),
            output_file=str(output_path),
            success=True,
            sheet_count=2,
        )

    monkeypatch.setattr("xlsliberator.web.runner.convert", fake_convert)

    WebJobRunner(store, WebSettings(data_dir=tmp_path)).run_job(job_id)
    job = store.get_job(job_id)

    assert job is not None
    assert job.status == JobPhase.COMPLETED
    assert job.report_json_path.exists()
    assert str(tmp_path) not in job.report_json_path.read_text()
    assert calls["allow_global_macro_security_change"] is False
    assert calls["user_installation_dir"] == tmp_path / "profile"


def test_runner_failure_marks_job_failed(tmp_path: Path, monkeypatch: Any) -> None:
    store, job_id = _store_with_job(tmp_path)

    def fake_convert(*_args: Any, **_kwargs: Any) -> ConversionReport:
        raise RuntimeError("boom")

    monkeypatch.setattr("xlsliberator.web.runner.convert", fake_convert)

    WebJobRunner(store, WebSettings(data_dir=tmp_path)).run_job(job_id)
    job = store.get_job(job_id)

    assert job is not None
    assert job.status == JobPhase.FAILED
    assert job.error == "boom"
