#!/usr/bin/env python3
"""Wrapper to run xlsliberator with UNO support."""

import os
import sys
from pathlib import Path

# Set UNO environment variables
os.environ["URE_BOOTSTRAP"] = "file:///usr/lib/libreoffice/program/fundamentalrc"
os.environ["LD_LIBRARY_PATH"] = "/usr/lib/x86_64-linux-gnu:/usr/lib/libreoffice/program"

# Activate virtual environment packages first
venv_site = Path(__file__).parent / ".venv/lib/python3.12/site-packages"
if venv_site.exists():
    sys.path.insert(0, str(venv_site))

# Add xlsliberator src to path
src_path = Path(__file__).parent / "src"
sys.path.insert(0, str(src_path))

# Add UNO path at the END so our packages take precedence
sys.path.append("/usr/lib/python3/dist-packages")

# Now run the CLI
from xlsliberator.cli import cli  # noqa: E402

if __name__ == "__main__":
    cli()
