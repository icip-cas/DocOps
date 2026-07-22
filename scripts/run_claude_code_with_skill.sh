#!/usr/bin/env zsh
set -euo pipefail
ROOT=$(cd -- "$(dirname -- "$0")/.." && pwd)
exec "$ROOT/scripts/run_harness.sh" claude-code --skill on "$@"

