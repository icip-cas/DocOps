#!/usr/bin/env zsh
set -euo pipefail
ROOT=$(cd -- "$(dirname -- "$0")/.." && pwd)
exec "$ROOT/docker/load_images.sh" "${1:-x86}"
