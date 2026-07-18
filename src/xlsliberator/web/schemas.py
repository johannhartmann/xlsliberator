"""Web application settings and shared schema helpers."""

from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class WebSettings:
    """Configuration for the XLSLiberator web app."""

    data_dir: Path
    max_upload_mb: int = 64
    worker_count: int = 1
    job_retention_hours: int = 24
    embed_macros: bool = False
    use_agent: bool = False
    open_swe_url: str = ""
    open_swe_token: str = ""
    open_swe_owner_id: str = "xlsliberator-web"
    open_swe_poll_seconds: float = 1.0
    open_swe_request_timeout_seconds: float = 60.0
    open_swe_job_timeout_seconds: int = 3600

    @classmethod
    def from_env(cls) -> WebSettings:
        """Build settings from environment variables."""
        return cls(
            data_dir=Path(os.getenv("XLSLIBERATOR_DATA_DIR", "/data")),
            max_upload_mb=_int_env("XLSLIBERATOR_MAX_UPLOAD_MB", 64),
            worker_count=max(1, _int_env("XLSLIBERATOR_WEB_WORKERS", 1)),
            job_retention_hours=max(1, _int_env("XLSLIBERATOR_JOB_RETENTION_HOURS", 24)),
            embed_macros=os.getenv("XLSLIBERATOR_EMBED_MACROS", "0") == "1",
            use_agent=os.getenv("XLSLIBERATOR_USE_AGENT", "0") == "1",
            open_swe_url=os.getenv("XLSLIBERATOR_OPEN_SWE_URL", ""),
            open_swe_token=os.getenv("XLSLIBERATOR_OPEN_SWE_TOKEN", ""),
            open_swe_owner_id=os.getenv(
                "XLSLIBERATOR_OPEN_SWE_OWNER_ID", "xlsliberator-web"
            ),
            open_swe_poll_seconds=max(
                0.1, _float_env("XLSLIBERATOR_OPEN_SWE_POLL_SECONDS", 1.0)
            ),
            open_swe_request_timeout_seconds=max(
                1.0,
                _float_env("XLSLIBERATOR_OPEN_SWE_REQUEST_TIMEOUT_SECONDS", 60.0),
            ),
            open_swe_job_timeout_seconds=max(
                30, _int_env("XLSLIBERATOR_OPEN_SWE_JOB_TIMEOUT_SECONDS", 3600)
            ),
        )


def _int_env(name: str, default: int) -> int:
    raw = os.getenv(name)
    if raw is None:
        return default
    try:
        return int(raw)
    except ValueError:
        return default


def _float_env(name: str, default: float) -> float:
    raw = os.getenv(name)
    if raw is None:
        return default
    try:
        return float(raw)
    except ValueError:
        return default
