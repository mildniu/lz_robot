#!/usr/bin/env bash

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

VENV_DIR=".venv311"
PYTHON_BIN="$VENV_DIR/bin/python3"

echo "============================================"
echo "Quantum Bot - Tk Desktop App v5.2 (macOS)"
echo "============================================"
echo

if ! command -v python3 >/dev/null 2>&1; then
  echo "[ERROR] python3 not found. Please install Python 3 first."
  exit 1
fi

if [ ! -x "$PYTHON_BIN" ]; then
  echo "[1/4] Creating virtual environment..."
  python3 -m venv "$VENV_DIR"
fi

echo "[2/4] Ensuring base packaging tools..."
"$PYTHON_BIN" -m pip install --upgrade pip setuptools >/dev/null

echo "[3/4] Installing Tk app dependencies if needed..."
"$PYTHON_BIN" - <<'PY'
import importlib
import subprocess
import sys

required = {
    "customtkinter": "customtkinter==5.2.0",
    "watchdog": "watchdog==3.0.0",
    "pystray": "pystray==0.19.5",
    "requests": "requests>=2.31.0",
}

missing = []
for module_name, package_spec in required.items():
    try:
        importlib.import_module(module_name)
    except Exception:
        missing.append(package_spec)

if missing:
    subprocess.check_call([sys.executable, "-m", "pip", "install", *missing])
PY

echo "[4/4] Starting Tk desktop application..."
echo
"$PYTHON_BIN" gui_app.py
