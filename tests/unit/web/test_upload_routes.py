from pathlib import Path
from typing import Any

from fastapi.testclient import TestClient

from xlsliberator.web.app import create_app
from xlsliberator.web.schemas import WebSettings


def _client(tmp_path: Path, monkeypatch: Any) -> TestClient:
    app = create_app(WebSettings(data_dir=tmp_path))
    monkeypatch.setattr(app.state.job_runner, "submit", lambda _job_id: None)
    return TestClient(app)


def test_valid_upload_creates_job_and_file(tmp_path: Path, monkeypatch: Any) -> None:
    client = _client(tmp_path, monkeypatch)

    response = client.post(
        "/api/jobs",
        files={"file": ("book.xlsx", b"PK\x03\x04content", "application/octet-stream")},
    )

    assert response.status_code == 202
    payload = response.json()
    assert payload["status"] == "queued"
    job_id = payload["id"]
    assert (tmp_path / "jobs" / job_id / "input.xlsx").exists()
    assert str(tmp_path) not in str(payload)


def test_browser_upload_redirects(tmp_path: Path, monkeypatch: Any) -> None:
    client = _client(tmp_path, monkeypatch)

    response = client.post(
        "/jobs",
        files={"file": ("book.xlsx", b"PK\x03\x04content", "application/octet-stream")},
        follow_redirects=False,
    )

    assert response.status_code == 303
    assert response.headers["location"].startswith("/jobs/")


def test_invalid_upload_extension_fails(tmp_path: Path, monkeypatch: Any) -> None:
    client = _client(tmp_path, monkeypatch)

    response = client.post("/api/jobs", files={"file": ("book.txt", b"text")})

    assert response.status_code == 400


def test_unknown_job_returns_404(tmp_path: Path, monkeypatch: Any) -> None:
    client = _client(tmp_path, monkeypatch)

    response = client.get("/api/jobs/unknown")

    assert response.status_code == 404
