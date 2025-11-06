"""Pytest configuration and shared fixtures."""

import os
import sys

import pytest

# Add system Python path for UNO modules (must be at beginning to access UNO)
# Place it at the end to avoid conflicts with our dependencies
if "/usr/lib/python3/dist-packages" not in sys.path:
    sys.path.append("/usr/lib/python3/dist-packages")


@pytest.fixture
def skip_if_no_lo() -> None:
    """Skip test if LibreOffice is not available."""
    if os.getenv("LO_SKIP_IT") == "1":
        pytest.skip("LibreOffice integration tests disabled (LO_SKIP_IT=1)")
