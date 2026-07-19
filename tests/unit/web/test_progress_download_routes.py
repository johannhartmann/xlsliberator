import json
from pathlib import Path
from typing import Any

from fastapi.testclient import TestClient

from xlsliberator.report import ConversionReport
from xlsliberator.web.app import create_app
from xlsliberator.web.jobs import JobArtifact, JobPhase
from xlsliberator.web.schemas import WebSettings


def _completed_client(tmp_path: Path) -> tuple[TestClient, str]:
    app = create_app(WebSettings(data_dir=tmp_path))
    store = app.state.job_store
    job_id = "33333333-3333-4333-8333-333333333333"
    job_dir = tmp_path / "jobs" / job_id
    job_dir.mkdir(parents=True)
    output = job_dir / "output.ods"
    report_json = job_dir / "report.json"
    report_md = job_dir / "report.md"
    output.write_text("ods")
    report = ConversionReport(
        input_file="book.xlsx",
        output_file="output.ods",
        success=True,
        sheet_count=3,
        total_cells=10,
        total_formulas=4,
        formula_match_rate=100,
        warnings=["minor"],
    )
    report.save_json(report_json)
    report.save_markdown(report_md)
    store.create_job(
        job_id=job_id,
        original_filename="book.xlsx",
        input_path=job_dir / "input.xlsx",
        output_path=output,
        report_json_path=report_json,
        report_md_path=report_md,
        log_bundle_path=job_dir / "logs.zip",
        profile_dir=job_dir / "profile",
    )
    store.mark_completed(job_id)
    return TestClient(app), job_id


def test_events_endpoint_returns_ordered_events(tmp_path: Path) -> None:
    client, job_id = _completed_client(tmp_path)

    response = client.get(f"/api/jobs/{job_id}/events?since=0")

    assert response.status_code == 200
    payload = response.json()
    assert payload["events"][0]["index"] == 0
    assert payload["next"] == len(payload["events"])


def test_report_api_returns_summary_for_completed_job(tmp_path: Path) -> None:
    client, job_id = _completed_client(tmp_path)

    response = client.get(f"/api/jobs/{job_id}/report")

    assert response.status_code == 200
    summary = response.json()
    assert summary["sheet_count"] == 3
    assert summary["total_formulas"] == 4
    assert summary["formula_match_rate"] == 100
    assert summary["warnings"] == ["minor"]


def test_report_api_404_when_no_report(tmp_path: Path) -> None:
    app = create_app(WebSettings(data_dir=tmp_path))
    store = app.state.job_store
    job_id = "55555555-5555-4555-8555-555555555555"
    job_dir = tmp_path / "jobs" / job_id
    job_dir.mkdir(parents=True)
    store.create_job(
        job_id=job_id,
        original_filename="book.xlsx",
        input_path=job_dir / "input.xlsx",
        output_path=job_dir / "output.ods",
        report_json_path=job_dir / "missing-report.json",
        report_md_path=job_dir / "report.md",
        log_bundle_path=job_dir / "logs.zip",
        profile_dir=job_dir / "profile",
    )
    store.add_event(job_id, phase=JobPhase.QUEUED, step="queued", message="queued")

    assert TestClient(app).get(f"/api/jobs/{job_id}/report").status_code == 404


def test_downloads_for_completed_job(tmp_path: Path) -> None:
    client, job_id = _completed_client(tmp_path)

    assert client.get(f"/jobs/{job_id}/download").status_code == 200
    assert client.get(f"/jobs/{job_id}/report.json").status_code == 200
    assert client.get(f"/jobs/{job_id}/report.md").status_code == 200


def test_download_rejects_completed_phase_without_operation_pass(tmp_path: Path) -> None:
    app = create_app(WebSettings(data_dir=tmp_path))
    store = app.state.job_store
    job_id = "66666666-6666-4666-8666-666666666666"
    job_dir = tmp_path / "jobs" / job_id
    job_dir.mkdir(parents=True)
    store.create_job(
        job_id=job_id,
        original_filename="book.xlsx",
        input_path=job_dir / "input.xlsx",
        output_path=job_dir / "output.ods",
        report_json_path=job_dir / "report.json",
        report_md_path=job_dir / "report.md",
        log_bundle_path=job_dir / "logs.zip",
        profile_dir=job_dir / "profile",
    )
    store.add_event(
        job_id,
        phase=JobPhase.COMPLETED,
        step="completed",
        message="Unverified progress event",
    )

    assert TestClient(app).get(f"/jobs/{job_id}/download").status_code == 409


def test_incomplete_job_cannot_download(tmp_path: Path) -> None:
    app = create_app(WebSettings(data_dir=tmp_path))
    store = app.state.job_store
    job_id = "44444444-4444-4444-8444-444444444444"
    job_dir = tmp_path / "jobs" / job_id
    job_dir.mkdir(parents=True)
    store.create_job(
        job_id=job_id,
        original_filename="book.xlsx",
        input_path=job_dir / "input.xlsx",
        output_path=job_dir / "output.ods",
        report_json_path=job_dir / "report.json",
        report_md_path=job_dir / "report.md",
        log_bundle_path=job_dir / "logs.zip",
        profile_dir=job_dir / "profile",
    )
    store.add_event(job_id, phase=JobPhase.QUEUED, step="queued", message="queued")

    assert TestClient(app).get(f"/jobs/{job_id}/download").status_code == 409


def test_completed_page_displays_report_metrics(tmp_path: Path) -> None:
    client, job_id = _completed_client(tmp_path)

    response = client.get(f"/jobs/{job_id}")

    assert response.status_code == 200
    assert "Formula match rate" in response.text
    assert "100.00%" in response.text


def test_completed_showcase_replays_real_job_artifacts(tmp_path: Path) -> None:
    client, job_id = _completed_client(tmp_path)
    store = client.app.state.job_store
    job = store.get_job(job_id)
    assert job is not None
    recording = job.input_path.parent / "keyboard-control.webm"
    recording.write_bytes(b"real-webm-evidence")
    result = job.input_path.parent / "keyboard-control-result.json"
    result.write_text(
        json.dumps(
            {
                "status": "passed",
                "scenario_id": "keyboard-control",
                "operations": [
                    {
                        "sequence": 1,
                        "kind": "key",
                        "status": "passed",
                        "duration_ms": 125,
                    }
                ],
            }
        ),
        encoding="utf-8",
    )
    store.set_artifacts(
        job_id,
        [
            JobArtifact(
                id="a" * 24,
                name=recording.name,
                kind="showcase-recording",
                media_type="video/webm",
                path=recording,
            ),
            JobArtifact(
                id="b" * 24,
                name=result.name,
                kind="showcase-result",
                media_type="application/json",
                path=result,
            ),
        ],
    )

    response = client.get(f"/jobs/{job_id}/showcase")

    assert response.status_code == 200
    assert "Application replay · keyboard-control" in response.text
    assert f"/jobs/{job_id}/artifacts/{'a' * 24}" in response.text
    assert "operation.duration_ms" in response.text
    assert client.get(f"/api/jobs/{job_id}").json()["downloads"]["showcase_replay"] == (
        f"/jobs/{job_id}/showcase"
    )

    result.write_text('{"status":"failed"}', encoding="utf-8")
    rejected = client.get(f"/jobs/{job_id}/showcase")
    assert rejected.status_code == 409
    assert rejected.json()["detail"] == "Showcase result did not pass"


def test_cancel_api_for_queued_job(tmp_path: Path, monkeypatch: Any) -> None:
    app = create_app(WebSettings(data_dir=tmp_path))
    monkeypatch.setattr(app.state.job_runner, "submit", lambda _job_id: None)
    client = TestClient(app)
    upload = client.post(
        "/api/jobs",
        files={"file": ("book.xlsx", b"PK\x03\x04content", "application/octet-stream")},
    )
    job_id = upload.json()["id"]

    response = client.post(f"/api/jobs/{job_id}/cancel")

    assert response.status_code == 200
    assert response.json()["cancellation_requested"] is True
