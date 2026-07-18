"""Authenticated client for the Open-SWE workbook migration API."""

from __future__ import annotations

import base64
import hashlib
import json
import re
import uuid
from pathlib import Path
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.parse import urlsplit
from urllib.request import Request, urlopen

LIBREOFFICE_BUILD = "26.2.4.2"
_MAX_RESPONSE_BYTES = 64 * 1024 * 1024
_ARTIFACT_ID = re.compile(r"^[0-9a-f]{24}$")


class OpenSWEError(RuntimeError):
    """A safe, user-facing Open-SWE transport or operation failure."""


class OpenSWEClient:
    """Small synchronous SDK used by the bounded web worker pool."""

    def __init__(
        self,
        *,
        base_url: str,
        token: str,
        owner_id: str,
        timeout_seconds: float = 60.0,
    ) -> None:
        if not base_url.strip() or not token.strip() or not owner_id.strip():
            raise ValueError("Open-SWE URL, token and owner ID are required")
        clean_base_url = base_url.strip()
        clean_token = token.strip()
        clean_owner_id = owner_id.strip()
        parsed = urlsplit(clean_base_url)
        if (
            parsed.scheme not in {"http", "https"}
            or not parsed.netloc
            or parsed.username is not None
            or parsed.password is not None
            or parsed.path not in {"", "/"}
            or parsed.query
            or parsed.fragment
        ):
            raise ValueError("Open-SWE URL must be an HTTP(S) service origin")
        if any(
            ord(character) < 32 or ord(character) == 127
            for character in clean_token + clean_owner_id
        ):
            raise ValueError("Open-SWE token and owner ID must not contain control characters")
        self.base_url = clean_base_url.rstrip("/")
        self.token = clean_token
        self.owner_id = clean_owner_id
        self.timeout_seconds = timeout_seconds

    def create_migration(self, workbook: Path, requirements: str = "") -> dict[str, Any]:
        data = workbook.read_bytes()
        payload = {
            "owner_id": self.owner_id,
            "artifact": {
                "original_filename": workbook.name,
                "sha256": hashlib.sha256(data).hexdigest(),
                "media_type": _workbook_media_type(workbook),
                "artifact_base64": base64.b64encode(data).decode(),
            },
            "user_requirements": requirements,
            "target_libreoffice_version": LIBREOFFICE_BUILD,
            "privacy_retention": {
                "classification": "private",
                "retain_days": 14,
                "delete_source_after_completion": True,
            },
        }
        result = self._json_request("POST", "/api/xlsliberator/migrations", payload)
        return _object(result)

    def status(self, thread_id: str) -> dict[str, Any]:
        return _object(self._json_request("GET", self._migration_path(thread_id)))

    def events(self, thread_id: str, since: int) -> dict[str, Any]:
        path = f"{self._migration_path(thread_id)}/events?since={max(0, since)}"
        return _object(self._json_request("GET", path))

    def follow_up(
        self,
        thread_id: str,
        *,
        requirements: str = "",
        dependency: Path | None = None,
        media_type: str = "application/octet-stream",
    ) -> dict[str, Any]:
        payload: dict[str, Any] = {"requirements": requirements}
        if dependency is not None:
            data = dependency.read_bytes()
            payload["dependency"] = {
                "original_filename": dependency.name,
                "sha256": hashlib.sha256(data).hexdigest(),
                "media_type": media_type,
                "artifact_base64": base64.b64encode(data).decode(),
            }
        return _object(
            self._json_request(
                "POST",
                f"{self._migration_path(thread_id)}/follow-ups",
                payload,
            )
        )

    def cancel(self, thread_id: str) -> dict[str, Any]:
        return _object(self._json_request("POST", f"{self._migration_path(thread_id)}/cancel", {}))

    def cleanup(self, thread_id: str) -> dict[str, Any]:
        """Delete the private Open-SWE migration workspace."""
        return _object(self._json_request("DELETE", self._migration_path(thread_id)))

    def download_artifact(self, thread_id: str, artifact_id: str) -> bytes:
        if _ARTIFACT_ID.fullmatch(artifact_id) is None:
            raise OpenSWEError("Invalid migration artifact identifier")
        path = f"{self._migration_path(thread_id)}/artifacts/{artifact_id}"
        return self._request("GET", path, None, expect_json=False)

    def ready(self) -> bool:
        try:
            self._json_request("GET", "/health")
        except OpenSWEError:
            return False
        return True

    def _migration_path(self, thread_id: str) -> str:
        try:
            parsed = uuid.UUID(thread_id)
        except ValueError:
            raise OpenSWEError("Invalid migration thread identifier") from None
        canonical = str(parsed)
        if thread_id != canonical:
            raise OpenSWEError("Invalid migration thread identifier")
        return f"/api/xlsliberator/migrations/{canonical}"

    def _json_request(
        self,
        method: str,
        path: str,
        payload: dict[str, Any] | None = None,
    ) -> object:
        raw = self._request(method, path, payload, expect_json=True)
        try:
            return json.loads(raw)
        except (TypeError, json.JSONDecodeError) as exc:
            raise OpenSWEError("Open-SWE returned an invalid response") from exc

    def _request(
        self,
        method: str,
        path: str,
        payload: dict[str, Any] | None,
        *,
        expect_json: bool,
    ) -> bytes:
        body = json.dumps(payload).encode() if payload is not None else None
        headers = {
            "Authorization": f"Bearer {self.token}",
            "X-XLSLiberator-Owner": self.owner_id,
            "Accept": "application/json" if expect_json else "application/octet-stream",
        }
        if body is not None:
            headers["Content-Type"] = "application/json"
        request = Request(
            f"{self.base_url}{path}",
            data=body,
            headers=headers,
            method=method,
        )
        try:
            with urlopen(request, timeout=self.timeout_seconds) as response:
                content = response.read(_MAX_RESPONSE_BYTES + 1)
                if len(content) > _MAX_RESPONSE_BYTES:
                    raise OpenSWEError("Open-SWE response exceeded the safety limit")
                return bytes(content)
        except HTTPError as exc:
            detail = _safe_http_error(exc)
            raise OpenSWEError(detail) from exc
        except (OSError, URLError) as exc:
            raise OpenSWEError("Open-SWE migration service is unavailable") from exc


def _object(value: object) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise OpenSWEError("Open-SWE returned an invalid response")
    return value


def _safe_http_error(error: HTTPError) -> str:
    if error.code == 401:
        return "Open-SWE authentication failed"
    if error.code == 403:
        return "Open-SWE rejected this migration operation"
    if error.code == 404:
        return "Open-SWE migration was not found"
    if error.code == 409:
        return "Open-SWE migration is not ready for this operation"
    if error.code == 422:
        return "Open-SWE rejected the supplied workbook or dependency"
    return f"Open-SWE migration service returned HTTP {error.code}"


def _workbook_media_type(path: Path) -> str:
    return {
        ".xls": "application/vnd.ms-excel",
        ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        ".xlsm": "application/vnd.ms-excel.sheet.macroenabled.12",
        ".xlsb": "application/vnd.ms-excel.sheet.binary.macroenabled.12",
    }.get(path.suffix.lower(), "application/octet-stream")
