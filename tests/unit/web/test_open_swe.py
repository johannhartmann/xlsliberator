from __future__ import annotations

import json
from pathlib import Path
from typing import Any
from urllib.request import Request

import pytest

from xlsliberator.web import open_swe
from xlsliberator.web.open_swe import OpenSWEClient, OpenSWEError


class _Response:
    def __init__(self, content: bytes) -> None:
        self.content = content

    def __enter__(self) -> _Response:
        return self

    def __exit__(self, *_args: object) -> None:
        return None

    def read(self, limit: int) -> bytes:
        return self.content[:limit]


def test_create_migration_sends_authenticated_private_request(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = tmp_path / "book.xlsx"
    workbook.write_bytes(b"PK\x03\x04workbook")
    captured: dict[str, Any] = {}

    def urlopen(request: Request, *, timeout: float) -> _Response:
        captured["request"] = request
        captured["timeout"] = timeout
        return _Response(b'{"thread_id":"aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa"}')

    monkeypatch.setattr(open_swe, "urlopen", urlopen)
    client = OpenSWEClient(
        base_url="https://swe.example.test/",
        token="secret-token",
        owner_id="tenant-1",
        timeout_seconds=12.0,
    )

    result = client.create_migration(workbook)

    assert result["thread_id"] == "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa"
    request = captured["request"]
    assert isinstance(request, Request)
    assert request.full_url == "https://swe.example.test/api/xlsliberator/migrations"
    headers = dict(request.header_items())
    assert headers["Authorization"] == "Bearer secret-token"
    assert headers["X-xlsliberator-owner"] == "tenant-1"
    payload = json.loads(request.data or b"{}")
    assert payload["owner_id"] == "tenant-1"
    assert payload["target_libreoffice_version"] == "26.2.4.2"
    assert payload["privacy_retention"]["delete_source_after_completion"] is True
    assert captured["timeout"] == 12.0


def test_download_rejects_non_opaque_artifact_identifier() -> None:
    client = OpenSWEClient(
        base_url="https://swe.example.test",
        token="secret-token",
        owner_id="tenant-1",
    )

    with pytest.raises(OpenSWEError, match="artifact identifier"):
        client.download_artifact(
            "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa",
            "../private",
        )


def test_status_rejects_noncanonical_thread_identifier() -> None:
    client = OpenSWEClient(
        base_url="https://swe.example.test",
        token="secret-token",
        owner_id="tenant-1",
    )

    with pytest.raises(OpenSWEError, match="thread identifier"):
        client.status("not-a-thread")
