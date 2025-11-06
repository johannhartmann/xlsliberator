#!/bin/bash
# Wrapper to run xlsliberator with UNO support

# Ensure LibreOffice is running
if ! nc -z 127.0.0.1 2002 2>/dev/null; then
    echo "Starting LibreOffice headless..."
    soffice --headless --accept="socket,host=127.0.0.1,port=2002;urp;" > /dev/null 2>&1 &
    sleep 3
fi

# Force system libstdc++ for LibreOffice UNO compatibility
export LD_PRELOAD="/usr/lib/x86_64-linux-gnu/libstdc++.so.6"
export LD_LIBRARY_PATH="/usr/lib/x86_64-linux-gnu:/usr/lib/libreoffice/program"
export URE_BOOTSTRAP="file:///usr/lib/libreoffice/program/fundamentalrc"

# Run xlsliberator - run_xlsliberator.py handles UNO path internally
exec python run_xlsliberator.py "$@"
