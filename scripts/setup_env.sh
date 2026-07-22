#!/usr/bin/env zsh
set -euo pipefail

ROOT=$(cd -- "$(dirname -- "$0")/.." && pwd)
PYTHON_BIN="${PYTHON_BIN:-python3.12}"

cd "$ROOT"

"$PYTHON_BIN" -c 'import sys; raise SystemExit(0 if sys.version_info >= (3, 12) else "Python 3.12+ is required by Harbor.")'
"$PYTHON_BIN" -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

echo "Environment ready: $ROOT/.venv"
