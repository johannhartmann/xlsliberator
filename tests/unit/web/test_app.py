from pathlib import Path

from fastapi.testclient import TestClient

from xlsliberator.web.app import create_app
from xlsliberator.web.jobs import JobPhase, JobStore
from xlsliberator.web.schemas import WebSettings


def test_app_health_and_readyz_without_open_swe_configuration(tmp_path: Path) -> None:
    app = create_app(WebSettings(data_dir=tmp_path))
    client = TestClient(app)

    assert client.get("/healthz").json() == {"status": "ok"}
    ready = client.get("/readyz").json()
    assert ready["data_dir_writable"] is True
    assert ready["open_swe_configured"] is False
    assert ready["open_swe_reachable"] is False
    assert ready["target_libreoffice_version"] == "26.2.4.2"


def test_index_page_renders_marketing_landing_and_demo(tmp_path: Path) -> None:
    client = TestClient(create_app(WebSettings(data_dir=tmp_path)))

    response = client.get("/")

    assert response.status_code == 200
    # Marketing content is present alongside the live demo.
    assert "XLSLiberator" in response.text
    assert "Digitale Souveränität" in response.text
    # The demo widget and its real upload form are wired up.
    assert 'id="demo-form"' in response.text
    assert 'action="/jobs"' in response.text
    assert "/static/demo.js" in response.text
    script = client.get("/static/demo.js")
    assert script.status_code == 200
    assert "LEVEL 0 · BASIS-PIPELINE" in script.text


def test_startup_resumes_persisted_open_swe_thread(tmp_path: Path) -> None:
    job_id = "88888888-8888-4888-8888-888888888888"
    job_dir = tmp_path / "jobs" / job_id
    job_dir.mkdir(parents=True)
    store = JobStore(tmp_path)
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
    store.add_event(job_id, phase=JobPhase.ANALYZING, step="lead", message="running")
    store.set_remote(
        job_id,
        thread_id="aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa",
        run_id="run-1",
    )
    app = create_app(WebSettings(data_dir=tmp_path))
    resumed: list[str] = []
    app.state.job_runner.resume = resumed.append

    with TestClient(app):
        pass

    assert resumed == [job_id]
