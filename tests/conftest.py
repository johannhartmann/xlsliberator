"""Pytest configuration and shared fixtures."""

import os
import shutil
import subprocess

import pytest

_UNEXPECTED_SKIPS: list[str] = []

# These files exercised the removed in-process host UNO API. Keeping them as
# historical test sources is useful while their remaining assertions are moved
# to worker-protocol tests, but importing or collecting them would violate the
# Docker-only runtime boundary.
collect_ignore = [
    "it/test_agent_runtime_validation.py",
    "it/test_macro_embed.py",
    "it/test_ods_writer_smoke.py",
    "it/test_translated_macro_runs.py",
    "it/test_uno_conn.py",
]


def pytest_runtest_logreport(report: pytest.TestReport) -> None:
    """Record skips in fail-closed CI jobs unless explicitly marked optional."""
    if (
        os.environ.get("XLSLIBERATOR_FAIL_ON_SKIP") == "1"
        and report.skipped
        and "optional" not in report.keywords
    ):
        _UNEXPECTED_SKIPS.append(report.nodeid)


def pytest_sessionfinish(session: pytest.Session) -> None:
    """Make an unexpected required skip fail the test process."""
    if _UNEXPECTED_SKIPS:
        session.exitstatus = pytest.ExitCode.TESTS_FAILED


@pytest.fixture
def skip_if_no_lo() -> None:
    """Skip only when the authoritative Docker runtime is unavailable."""
    if shutil.which("docker") is None:
        pytest.skip("LibreOffice integration tests require Docker")
    result = subprocess.run(
        ["docker", "image", "inspect", "xlsliberator-libreoffice:26.2.4.2"],
        capture_output=True,
        check=False,
        timeout=10,
    )
    if result.returncode != 0:
        pytest.skip("Pinned LibreOffice Docker image is not built")
