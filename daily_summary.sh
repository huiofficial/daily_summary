#!/bin/bash
# Shell wrapper for daily_summary.py
# This script provides a convenient way to run the Python script

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PYTHON_SCRIPT="$SCRIPT_DIR/daily_summary.py"

# Check if Python script exists
if [ ! -f "$PYTHON_SCRIPT" ]; then
    echo "Error: daily_summary.py not found at $PYTHON_SCRIPT"
    exit 1
fi

# Run the Python script with all arguments passed through
python3 "$PYTHON_SCRIPT" "$@"

