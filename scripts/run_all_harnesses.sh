#!/usr/bin/env zsh
set -euo pipefail

ROOT=$(cd -- "$(dirname -- "$0")/.." && pwd)

"$ROOT/scripts/run_doctools.sh"
"$ROOT/scripts/run_terminus2.sh"
"$ROOT/scripts/run_codex_with_skill.sh"
"$ROOT/scripts/run_codex_no_skill.sh"
"$ROOT/scripts/run_claude_code_with_skill.sh"
"$ROOT/scripts/run_claude_code_no_skill.sh"

