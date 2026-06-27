"""Retention cleanup for web job artifacts."""

from __future__ import annotations

import shutil
from datetime import UTC, datetime, timedelta
from pathlib import Path


class CleanupSafetyError(ValueError):
    """Raised when cleanup would target an unsafe directory."""


def cleanup_old_jobs(data_dir: Path, older_than: timedelta) -> list[Path]:
    """Delete job directories older than the configured retention window."""
    _assert_safe_data_dir(data_dir)
    jobs_dir = data_dir / "jobs"
    if not jobs_dir.exists():
        return []
    cutoff = datetime.now(UTC) - older_than
    deleted: list[Path] = []
    for child in jobs_dir.iterdir():
        if not child.is_dir():
            continue
        modified = datetime.fromtimestamp(child.stat().st_mtime, tz=UTC)
        if modified < cutoff:
            shutil.rmtree(child)
            deleted.append(child)
    return deleted


def _assert_safe_data_dir(data_dir: Path) -> None:
    resolved = data_dir.resolve()
    if resolved in {Path("/"), Path.home().resolve(), Path.cwd().resolve()}:
        raise CleanupSafetyError(f"Refusing to clean unsafe data directory: {resolved}")
