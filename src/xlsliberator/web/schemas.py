"""Web application settings and shared schema helpers."""

from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class WebSettings:
    """Configuration for the XLSLiberator web app."""

    data_dir: Path
    max_upload_mb: int = 100
    worker_count: int = 1
    job_retention_hours: int = 24
    embed_macros: bool = False
    use_agent: bool = False

    @classmethod
    def from_env(cls) -> WebSettings:
        """Build settings from environment variables."""
        return cls(
            data_dir=Path(os.getenv("XLSLIBERATOR_DATA_DIR", "/data")),
            max_upload_mb=_int_env("XLSLIBERATOR_MAX_UPLOAD_MB", 100),
            worker_count=max(1, _int_env("XLSLIBERATOR_WEB_WORKERS", 1)),
            job_retention_hours=max(1, _int_env("XLSLIBERATOR_JOB_RETENTION_HOURS", 24)),
            embed_macros=os.getenv("XLSLIBERATOR_EMBED_MACROS", "0") == "1",
            use_agent=os.getenv("XLSLIBERATOR_USE_AGENT", "0") == "1",
        )


def _int_env(name: str, default: int) -> int:
    raw = os.getenv(name)
    if raw is None:
        return default
    try:
        return int(raw)
    except ValueError:
        return default
