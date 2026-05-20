#!/bin/bash
# Double-clickable launcher: creates a venv on first run, installs deps,
# then launches the app. Subsequent runs reuse the existing venv.
set -e

DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$DIR"

if [ ! -d ".venv" ]; then
    echo "First run — creating .venv and installing dependencies (this takes a minute)..."
    python3 -m venv .venv
    # shellcheck disable=SC1091
    source .venv/bin/activate
    pip install --upgrade pip
    pip install -r requirements.txt
else
    # shellcheck disable=SC1091
    source .venv/bin/activate
    # If a dep is missing (e.g. user wiped site-packages), reinstall.
    if ! python -c "import PyQt6, polars, pandas, xlsxwriter" 2>/dev/null; then
        echo "Dependencies missing or broken — reinstalling..."
        pip install --upgrade pip
        pip install -r requirements.txt
    fi
fi

python main.py
