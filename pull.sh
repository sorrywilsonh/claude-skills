#!/bin/bash
# sync.sh 的反向操作：把 GitHub 上的 claude-skills 拉回本機 ~/.claude/skills/。
# 執行前會先確認本 repo 沒有未推送的變動，並在覆蓋本機前備份一份，避免尚未
# 跑過 sync.sh 的本機修改被悄悄蓋掉。
set -euo pipefail

REPO="$HOME/projects/claude-skills"
SRC="$REPO/.claude/skills/"
DEST="$HOME/.claude/skills/"
BACKUP_DIR="$HOME/.claude/skills-backups"

[ -d "$REPO/.git" ] || { echo "錯誤: $REPO 不是 git repo，先 git clone claude-skills 到這個路徑"; exit 1; }

cd "$REPO"

# repo 本身若有未 commit 的變動，代表可能還沒跑過 sync.sh，先擋下來避免被 git pull 弄丟
if [ -n "$(git status --porcelain)" ]; then
  echo "錯誤: $REPO 有未 commit 的變動，請先跑 sync.sh 推上去，或自行處理後再拉。"
  exit 1
fi

git fetch -q origin
# 用 --ff-only：本機 repo clone 若跟遠端分岔（例如手動改過又忘記推），直接失敗而不是自動合併
git pull --ff-only -q origin main

# 覆蓋本機 ~/.claude/skills 前先備份，防止尚未同步上去的本機修改被 --delete 一併清掉
mkdir -p "$BACKUP_DIR"
if [ -d "$DEST" ] && [ -n "$(ls -A "$DEST" 2>/dev/null)" ]; then
  cp -R "$DEST" "$BACKUP_DIR/$(date '+%Y%m%d-%H%M%S')"
fi

mkdir -p "$DEST"
rsync -a --delete "$SRC" "$DEST"

echo "已從 GitHub 拉回本機，備份於 $BACKUP_DIR，時間 $(date '+%Y-%m-%d %H:%M')"
