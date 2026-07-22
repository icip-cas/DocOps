#!/usr/bin/env zsh
set -euo pipefail

ROOT=$(cd -- "$(dirname -- "$0")/.." && pwd)
ENV_FILE="${DOCOPS_ENV_FILE:-$ROOT/.env}"

if [[ -f "$ENV_FILE" ]]; then
  set -a
  source "$ENV_FILE"
  set +a
fi

usage() {
  cat <<'EOF'
usage: scripts/run_harness.sh HARNESS [--skill on|off] [--output DIR] [--model MODEL] [-- EXTRA_HARBOR_ARGS...]

HARNESS:
  doctools
  terminus2
  codex
  claude-code

Examples:
  scripts/run_harness.sh doctools
  scripts/run_harness.sh terminus2 --model openai/gpt-5.5
  scripts/run_harness.sh codex --skill on
  scripts/run_harness.sh claude-code --skill off
EOF
}

if [[ $# -lt 1 || "${1:-}" == "-h" || "${1:-}" == "--help" ]]; then
  usage
  exit 0
fi

HARNESS="$1"
shift
SKILL="on"
MODEL_OVERRIDE=""
OUTPUT_OVERRIDE=""
EXTRA_ARGS=()

while [[ $# -gt 0 ]]; do
  case "$1" in
    --skill)
      SKILL="$2"
      shift 2
      ;;
    --model)
      MODEL_OVERRIDE="$2"
      shift 2
      ;;
    --output)
      OUTPUT_OVERRIDE="$2"
      shift 2
      ;;
    --)
      shift
      EXTRA_ARGS=("$@")
      break
      ;;
    -h|--help)
      usage
      exit 0
      ;;
    *)
      EXTRA_ARGS+=("$1")
      shift
      ;;
  esac
done

TASKS_DIR="${DOCOPS_TASKS_DIR:-$ROOT/tasks}"
RUN_TASKS_DIR="$TASKS_DIR"

if [[ "$SKILL" != "on" && "$SKILL" != "off" ]]; then
  echo "ERROR: --skill must be 'on' or 'off'." >&2
  exit 2
fi

if [[ "$SKILL" == "off" ]]; then
  RUN_TASKS_DIR="$ROOT/runtime/no_skill_tasks"
  "$ROOT/scripts/materialize_no_skill.py" "$TASKS_DIR" "$RUN_TASKS_DIR"
fi

HARBOR_REPO_ROOT="${HARBOR_REPO_ROOT:-$ROOT/third_party/harbor}"
if [[ ! -d "$HARBOR_REPO_ROOT/src/harbor" ]]; then
  echo "ERROR: Harbor source not found at $HARBOR_REPO_ROOT" >&2
  exit 1
fi

if [[ -z "${HARBOR_PYTHON_BIN:-}" ]]; then
  if [[ -x "$ROOT/.venv/bin/python" ]]; then
    HARBOR_PYTHON_BIN="$ROOT/.venv/bin/python"
  else
    HARBOR_PYTHON_BIN="python3"
  fi
elif [[ "$HARBOR_PYTHON_BIN" != /* ]]; then
  HARBOR_PYTHON_BIN="$ROOT/$HARBOR_PYTHON_BIN"
fi
HARBOR_PYTHONPATH="$ROOT/harnesses:$HARBOR_REPO_ROOT/src:$HARBOR_REPO_ROOT/packages/rewardkit/src"
if [[ -n "${PYTHONPATH:-}" ]]; then
  HARBOR_PYTHONPATH="$HARBOR_PYTHONPATH:$PYTHONPATH"
fi

RUN_COUNT="${HARBOR_RUN_COUNT:-1}"
N_CONCURRENT="${HARBOR_N_CONCURRENT:-1}"
OUTPUT_ROOT_BASE="${DOCOPS_OUTPUT_ROOT:-$ROOT/results}"
if [[ "$OUTPUT_ROOT_BASE" != /* ]]; then
  OUTPUT_ROOT_BASE="$ROOT/$OUTPUT_ROOT_BASE"
fi

build_common_cmd() {
  CMD=(
    env
    "PYTHONPATH=$HARBOR_PYTHONPATH"
  )
}

append_openai_env() {
  if [[ -n "${OPENAI_API_KEY:-}" ]]; then
    CMD+=("OPENAI_API_KEY=$OPENAI_API_KEY")
  fi
  if [[ -n "${OPENAI_BASE_URL:-}" ]]; then
    CMD+=("OPENAI_BASE_URL=$OPENAI_BASE_URL")
  fi
}

case "$HARNESS" in
  doctools)
    MODEL="${MODEL_OVERRIDE:-${DOCOPS_DOCTOOLS_MODEL:-openai/gpt-5.5}}"
    OUTPUT_ROOT="${OUTPUT_OVERRIDE:-$OUTPUT_ROOT_BASE/doctools}"
    if [[ -z "${OPENAI_API_KEY:-}" ]]; then
      echo "ERROR: OPENAI_API_KEY is required for DocumentTools." >&2
      exit 1
    fi
    build_common_cmd
    append_openai_env
    CMD+=(
      "$HARBOR_PYTHON_BIN" -m harbor.cli.main run
      -p "$RUN_TASKS_DIR"
      --agent-import-path document_tools_agent:DocumentToolsAgent
      -m "$MODEL"
      -o "$OUTPUT_ROOT"
      -n "$RUN_COUNT"
      --n-concurrent "$N_CONCURRENT"
      -y
      --ak "max_steps=${HARBOR_DT_MAX_STEPS:-12}"
      --ak "temperature=${HARBOR_DT_TEMPERATURE:-0.0}"
      --ak "reasoning_effort=${HARBOR_DT_REASONING_EFFORT:-high}"
    )
    if [[ -n "${OPENAI_BASE_URL:-}" ]]; then
      CMD+=(--ak "api_base=$OPENAI_BASE_URL")
    fi
    ;;
  terminus2)
    MODEL="${MODEL_OVERRIDE:-${DOCOPS_TERMINUS2_MODEL:-openai/gpt-5.5}}"
    OUTPUT_ROOT="${OUTPUT_OVERRIDE:-$OUTPUT_ROOT_BASE/terminus2}"
    if [[ -z "${OPENAI_API_KEY:-}" ]]; then
      echo "ERROR: OPENAI_API_KEY is required for Terminus-2." >&2
      exit 1
    fi
    build_common_cmd
    append_openai_env
    CMD+=(
      "$HARBOR_PYTHON_BIN" -m harbor.cli.main run
      -p "$RUN_TASKS_DIR"
      -a terminus-2
      -m "$MODEL"
      -o "$OUTPUT_ROOT"
      -n "$RUN_COUNT"
      --n-concurrent "$N_CONCURRENT"
      -y
      --ak "max_turns=${HARBOR_T2_MAX_TURNS:-200}"
      --ak "parser_name=${HARBOR_T2_PARSER:-json}"
      --ak "temperature=${HARBOR_T2_TEMPERATURE:-0.2}"
      --ak "reasoning_effort=\"${HARBOR_T2_REASONING_EFFORT:-high}\""
      --ak "enable_summarize=${HARBOR_T2_ENABLE_SUMMARIZE:-true}"
    )
    if [[ -n "${OPENAI_BASE_URL:-}" ]]; then
      CMD+=(--ak "api_base=$OPENAI_BASE_URL")
    fi
    ;;
  codex)
    MODEL="${MODEL_OVERRIDE:-${DOCOPS_CODEX_MODEL:-gpt-5.5}}"
    OUTPUT_ROOT="${OUTPUT_OVERRIDE:-$OUTPUT_ROOT_BASE/codex_${SKILL}_skill}"
    build_common_cmd
    append_openai_env
    CMD+=(
      "$HARBOR_PYTHON_BIN" -m harbor.cli.main run
      -p "$RUN_TASKS_DIR"
      -a codex
      -m "$MODEL"
      -o "$OUTPUT_ROOT"
      -n "$RUN_COUNT"
      --n-concurrent "$N_CONCURRENT"
      -y
      --ae "CODEX_FORCE_API_KEY=${CODEX_FORCE_API_KEY:-1}"
    )
    if [[ -n "${CODEX_CONFIG_TOML_PATH:-}" && -f "${CODEX_CONFIG_TOML_PATH:-}" ]]; then
      CMD+=(--ae "CODEX_CONFIG_TOML_PATH=$CODEX_CONFIG_TOML_PATH")
    fi
    if [[ -n "${OPENAI_BASE_URL:-}" ]]; then
      CMD+=(--ae "OPENAI_BASE_URL=$OPENAI_BASE_URL")
    fi
    if [[ -n "${OPENAI_API_KEY:-}" ]]; then
      CMD+=(--ae "OPENAI_API_KEY=$OPENAI_API_KEY")
    fi
    ;;
  claude-code)
    MODEL="${MODEL_OVERRIDE:-${DOCOPS_CLAUDE_MODEL:-claude-sonnet-4-6}}"
    OUTPUT_ROOT="${OUTPUT_OVERRIDE:-$OUTPUT_ROOT_BASE/claude_code_${SKILL}_skill}"
    if [[ -z "${ANTHROPIC_AUTH_TOKEN:-${ANTHROPIC_API_KEY:-}}" ]]; then
      echo "ERROR: ANTHROPIC_AUTH_TOKEN or ANTHROPIC_API_KEY is required for Claude Code." >&2
      exit 1
    fi
    if [[ -z "${ANTHROPIC_BASE_URL:-}" ]]; then
      echo "ERROR: ANTHROPIC_BASE_URL is required for Claude Code." >&2
      exit 1
    fi
    AUTH_TOKEN="${ANTHROPIC_AUTH_TOKEN:-$ANTHROPIC_API_KEY}"
    build_common_cmd
    CMD+=(
      "ANTHROPIC_BASE_URL=$ANTHROPIC_BASE_URL"
      "ANTHROPIC_AUTH_TOKEN=$AUTH_TOKEN"
      "ANTHROPIC_API_KEY=$AUTH_TOKEN"
      "$HARBOR_PYTHON_BIN" -m harbor.cli.main run
      -p "$RUN_TASKS_DIR"
      -a claude-code
      -m "$MODEL"
      -o "$OUTPUT_ROOT"
      -n "$RUN_COUNT"
      --n-concurrent "$N_CONCURRENT"
      -y
      --ae "ANTHROPIC_BASE_URL=$ANTHROPIC_BASE_URL"
      --ae "ANTHROPIC_AUTH_TOKEN=$AUTH_TOKEN"
      --ae "ANTHROPIC_API_KEY=$AUTH_TOKEN"
    )
    ;;
  *)
    echo "ERROR: unknown harness '$HARNESS'." >&2
    usage
    exit 2
    ;;
esac

echo "Harness: $HARNESS"
echo "Skill: $SKILL"
echo "Model: $MODEL"
echo "Tasks: $RUN_TASKS_DIR"
echo "Output: $OUTPUT_ROOT"

"${CMD[@]}" "${EXTRA_ARGS[@]}"
