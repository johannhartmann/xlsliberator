"""Fail-closed application boundary for the Docker-only platform."""

from __future__ import annotations

import os
from pathlib import Path

APPLICATION_CONTAINER_MARKER = "XLSLIBERATOR_APPLICATION_CONTAINER"
OFFICE_CONTAINER_MARKER = "XLSLIBERATOR_OFFICE_CONTAINER"


class ContainerBoundaryError(RuntimeError):
    """Raised when XLSLiberator application code is started on the host."""


def application_container_is_authorized() -> bool:
    """Return whether this process is an explicitly marked project container."""
    marked = os.environ.get(APPLICATION_CONTAINER_MARKER) == "1"
    office_worker = os.environ.get(OFFICE_CONTAINER_MARKER) == "1"
    return Path("/.dockerenv").is_file() and (marked or office_worker)


def require_application_container() -> None:
    """Reject host execution before conversion or a long-running service starts."""
    if not application_container_is_authorized():
        raise ContainerBoundaryError(
            "XLSLiberator is Docker-only. Host Python execution is forbidden; "
            "start the application through docker compose. LibreOffice and PyUNO "
            "must run only in the pinned LibreOffice 26.2.4.2 worker image."
        )
