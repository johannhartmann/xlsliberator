from pathlib import Path
from typing import Any

from fastapi.testclient import TestClient

from xlsliberator.web.app import create_app
from xlsliberator.web.schemas import WebSettings


def test_app_health_and_readyz_without_soffice(tmp_path: Path, monkeypatch: Any) -> None:
    monkeypatch.setattr("xlsliberator.web.app.shutil.which", lambda _name: None)
    app = create_app(WebSettings(data_dir=tmp_path))
    client = TestClient(app)

    assert client.get("/healthz").json() == {"status": "ok"}
    ready = client.get("/readyz").json()
    assert ready["data_dir_writable"] is True
    assert ready["soffice_available"] is False
    assert ready["version"] is None


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
