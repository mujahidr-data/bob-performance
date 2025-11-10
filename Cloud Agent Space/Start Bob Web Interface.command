#!/bin/bash
# Double-clickable launcher for HiBob Performance Report Web Interface

# Get the directory where this script is located
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
cd "$SCRIPT_DIR"

# Run the startup script
./scripts/shell/start_web_app.sh

# Keep terminal open if there's an error
if [ $? -ne 0 ]; then
    echo ""
    echo "Press Enter to close this window..."
    read
fi

