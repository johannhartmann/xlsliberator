#!/bin/bash
# Wrapper to run xlsliberator with UNO support

# Get script directory
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

# Ensure LibreOffice is running
if ! nc -z 127.0.0.1 2002 2>/dev/null; then
    echo "Starting LibreOffice headless..."
    soffice --headless --accept="socket,host=127.0.0.1,port=2002;urp;" > /dev/null 2>&1 &
    sleep 3
fi

# Activate virtual environment
source "$SCRIPT_DIR/.venv/bin/activate"

# CRITICAL: Set UNO environment BEFORE Python starts
# These must be set in the shell, not in Python, because uno bootstrap happens at import time
export URE_BOOTSTRAP="vnd.sun.star.pathname:/usr/lib/libreoffice/program/fundamentalrc"
export UNO_PATH="/usr/lib/libreoffice/program"
export LD_LIBRARY_PATH="/usr/lib/libreoffice/program:${LD_LIBRARY_PATH}"

# Run xlsliberator - run_xlsliberator.py handles path setup
exec python "$SCRIPT_DIR/run_xlsliberator.py" "$@"
