#!/bin/bash
# [推上雲端] 自動把 ~/.claude/skills 同步到本 repo 並推上 GitHub。
# 無變動時不做任何 commit/push（安靜結束）。反方向（拉回本機）見 pull.sh。
set -euo pipefail

SRC="$HOME/.claude/skills/"
REPO="$HOME/projects/claude-skills"
DEST="$REPO/.claude/skills/"

# 來源不存在就直接結束，避免把空目錄同步上去
[ -d "$SRC" ] || { echo "skip: $SRC 不存在"; exit 0; }

# 鏡像同步（含刪除）；git 歷史仍保留舊版本
rsync -a --delete "$SRC" "$DEST"

cd "$REPO"
if [ -z "$(git status --porcelain)" ]; then
  echo "no changes"
  exit 0
fi

git add -A
git commit -q -m "Auto-sync skills $(date '+%Y-%m-%d %H:%M')"
git push -q
echo "synced & pushed at $(date '+%Y-%m-%d %H:%M')"
