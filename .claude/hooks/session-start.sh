#!/bin/bash
set -euo pipefail

# 僅在雲端 session 執行
if [ "${CLAUDE_CODE_REMOTE:-}" != "true" ]; then
  exit 0
fi

REPO_URL="https://github.com/mvanhorn/last30days-skill"
CACHE_DIR="${HOME}/.cache/last30days-skill"
SKILL_DIR="${CLAUDE_PROJECT_DIR:-$(pwd)}/.claude/skills"

# Clone 或更新（idempotent）
if git -C "$CACHE_DIR" rev-parse HEAD >/dev/null 2>&1; then
  GIT_LFS_SKIP_SMUDGE=1 git -C "$CACHE_DIR" pull --ff-only --depth 1 -q || true
else
  rm -rf "$CACHE_DIR"
  GIT_LFS_SKIP_SMUDGE=1 git clone --depth 1 -q "$REPO_URL" "$CACHE_DIR"
fi

# 連結到專案 skills 目錄
mkdir -p "$SKILL_DIR"
ln -sfn "$CACHE_DIR/skills/last30days" "$SKILL_DIR/last30days"
