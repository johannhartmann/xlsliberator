"""Pytest configuration and shared fixtures."""

import os

import pytest


@pytest.fixture
def skip_if_no_lo() -> None:
    """Skip test if LibreOffice is not available."""
    if os.getenv("LO_SKIP_IT") == "1":
        pytest.skip("LibreOffice integration tests disabled (LO_SKIP_IT=1)")
