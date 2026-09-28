#!/usr/bin/env bash
# Exhibition Label Studio - set up profiles and settings (Linux / macOS)
cd "$(dirname "$0")" || exit 1

PY=python3
command -v python3 >/dev/null 2>&1 || PY=python
if ! command -v "$PY" >/dev/null 2>&1; then
    echo "Python 3 is not installed. Install it (e.g. 'sudo apt install python3 python3-venv') and try again."
    exit 1
fi

if [ ! -f venv/bin/activate ]; then
    echo "[SETUP] First run - creating a private Python environment..."
    "$PY" -m venv venv || { echo "Could not create the environment. Try: sudo apt install python3-venv"; exit 1; }
fi
# shellcheck disable=SC1091
source venv/bin/activate

if ! python -c "import openpyxl, reportlab" >/dev/null 2>&1; then
    echo "[SETUP] Installing required packages (one time, needs internet)..."
    python -m pip install --quiet --upgrade pip
    python -m pip install --quiet openpyxl reportlab || { echo "Package install failed - check your internet connection."; exit 1; }
fi

python Generate_Labels.py configure "$@"
