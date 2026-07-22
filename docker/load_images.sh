#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROFILE="${1:-x86}"
case "$PROFILE" in
  server-linux-amd64|linux-x86_64)
    PROFILE="x86"
    ;;
  linux-amd64|generic-linux-amd64)
    PROFILE="amd"
    ;;
esac
IMAGE_DIR="${SCRIPT_DIR}/images/${PROFILE}"

CLAUDE_CODE_VERSION="${CLAUDE_CODE_VERSION:-2.1.114}"
CODEX080_VERSION="${CODEX080_VERSION:-0.80.0}"
CLAUDE_IMAGE_NAME="${CLAUDE_IMAGE_NAME:-harbor-claude-code-base}"
CODEX_IMAGE_NAME="${CODEX_IMAGE_NAME:-harbor-codex-base}"
CODEX080_IMAGE_NAME="${CODEX080_IMAGE_NAME:-harbor-codex080-chat-base}"

if [[ ! -d "$IMAGE_DIR" ]]; then
  echo "ERROR: image profile not found: $PROFILE" >&2
  echo "Available profiles:" >&2
  find "${SCRIPT_DIR}/images" -mindepth 1 -maxdepth 1 -type d -exec basename {} \; | sort >&2
  exit 2
fi

shopt -s nullglob
CLAUDE_CANDIDATES=("${IMAGE_DIR}/harbor-claude-code-base_${CLAUDE_CODE_VERSION}_"*.tar.gz "${IMAGE_DIR}/harbor-claude-code-base_${CLAUDE_CODE_VERSION}_"*.tar)
CODEX_CANDIDATES=(
  "${IMAGE_DIR}/harbor-codex-base_${CLAUDE_CODE_VERSION}_"*.tar.gz
  "${IMAGE_DIR}/harbor-codex-base_${CLAUDE_CODE_VERSION}_"*.tar
  "${IMAGE_DIR}/harbor-codex080-chat-base_${CODEX080_VERSION}_"*.tar.gz
  "${IMAGE_DIR}/harbor-codex080-chat-base_${CODEX080_VERSION}_"*.tar
)

if [[ ${#CLAUDE_CANDIDATES[@]} -eq 0 || ${#CODEX_CANDIDATES[@]} -eq 0 ]]; then
  echo "ERROR: profile '$PROFILE' must contain both Harbor base image archives." >&2
  echo "Expected filename patterns:" >&2
  echo "  harbor-claude-code-base_${CLAUDE_CODE_VERSION}_*.tar.gz" >&2
  echo "  harbor-codex-base_${CLAUDE_CODE_VERSION}_*.tar.gz" >&2
  echo "  or harbor-codex080-chat-base_${CODEX080_VERSION}_*.tar.gz" >&2
  exit 1
fi

CLAUDE_TAR="${CLAUDE_CANDIDATES[0]}"
CODEX_TAR="${CODEX_CANDIDATES[0]}"

for image_tar in "$CLAUDE_TAR" "$CODEX_TAR"; do
  if [[ ! -f "$image_tar" ]]; then
    echo "ERROR: missing image archive: $image_tar" >&2
    exit 1
  fi
done

docker load -i "$CLAUDE_TAR"
docker load -i "$CODEX_TAR"

ensure_default_tag() {
  local image_name="$1"
  local default_tag="${image_name}:${CLAUDE_CODE_VERSION}"
  if docker image inspect "$default_tag" >/dev/null 2>&1; then
    return
  fi
  for suffix in amd64 arm64 linux-amd64 linux-arm64; do
    local candidate="${image_name}:${CLAUDE_CODE_VERSION}-${suffix}"
    if docker image inspect "$candidate" >/dev/null 2>&1; then
      docker tag "$candidate" "$default_tag"
      return
    fi
  done
  echo "ERROR: loaded image did not provide $default_tag or a known architecture-suffixed tag." >&2
  exit 1
}

ensure_default_tag "$CLAUDE_IMAGE_NAME"
if [[ "$(basename "$CODEX_TAR")" == harbor-codex080-chat-base_* ]]; then
  CODEX080_TAG="${CODEX080_IMAGE_NAME}:${CODEX080_VERSION}-amd64"
  if ! docker image inspect "$CODEX080_TAG" >/dev/null 2>&1; then
    CODEX080_TAG="${CODEX080_IMAGE_NAME}:${CODEX080_VERSION}"
  fi
  docker image inspect "$CODEX080_TAG" >/dev/null
  docker tag "$CODEX080_TAG" "${CODEX_IMAGE_NAME}:${CLAUDE_CODE_VERSION}"
else
  ensure_default_tag "$CODEX_IMAGE_NAME"
fi

echo "Loaded DocOps base images:"
echo "  ${CLAUDE_IMAGE_NAME}:${CLAUDE_CODE_VERSION}"
echo "  ${CODEX_IMAGE_NAME}:${CLAUDE_CODE_VERSION}"
