"""Load the repository's host-UNO import guard during Python startup."""

from __future__ import annotations

import runpy
from pathlib import Path

runpy.run_path(str(Path(__file__).resolve().parent / "src" / "sitecustomize.py"))
