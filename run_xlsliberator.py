#!/usr/bin/env python3
"""Wrapper to run xlsliberator with UNO support."""

import sys

# Add UNO path at the END so our packages take precedence
sys.path.append("/usr/lib/python3/dist-packages")

# Now run the CLI
from xlsliberator.cli import cli

if __name__ == "__main__":
    cli()
