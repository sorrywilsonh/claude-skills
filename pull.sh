#!/bin/bash
# sync.sh 的反向操作：把 GitHub 上的 claude-skills 拉回本機 ~/.claude/skills/。
# 執行前會先確認本 repo 沒有未推送的變動，並在覆蓋本機前備份一份，避免尚未
# 跑過 sync.sh 的本機修改被悄悄蓋掉。
#
# 全部包在 main() 裡：bash 會先把整個函式讀進記憶體再執行，這樣即使腳本在
# git pull 時把自己換成新版，正在執行的這一份也不會因檔案被改而錯亂
# （macOS 內建的舊 bash 3.2 對「執行中被改寫的腳本」特別脆弱）。
set -euo pipefail

main() {
  local REPO="$HOME/projects/claude-skills"
  local SRC="$REPO/.claude/skills/"
  local DEST="$HOME/.claude/skills/"
  local BACKUP_DIR="$HOME/.claude/skills-backups"

  [ -d "$REPO/.git" ] || { echo "錯誤: $REPO 不是 git repo，先 git clone claude-skills 到這個路徑"; exit 1; }

  cd "$REPO"

  # repo 本身若有未 commit 的變動，代表可能還沒跑過 sync.sh，先擋下來避免被 git pull 弄丟
  if [ -n "$(git status --porcelain)" ]; then
    echo "錯誤: $REPO 有未 commit 的變動，請先跑 sync.sh 推上去，或自行處理後再拉。"
    exit 1
  fi

  git fetch -q origin
  # 記下 pull 前後的 commit，稍後用來判斷有哪些 skill 的 requirements.txt 變動了
  local before_rev after_rev
  before_rev="$(git rev-parse HEAD)"
  # 用 --ff-only：本機 repo clone 若跟遠端分岔（例如手動改過又忘記推），直接失敗而不是自動合併
  git pull --ff-only -q origin main
  after_rev="$(git rev-parse HEAD)"

  # 覆蓋本機 ~/.claude/skills 前先備份，防止尚未同步上去的本機修改被 --delete 一併清掉
  mkdir -p "$BACKUP_DIR"
  if [ -d "$DEST" ] && [ -n "$(ls -A "$DEST" 2>/dev/null)" ]; then
    cp -R "$DEST" "$BACKUP_DIR/$(date '+%Y%m%d-%H%M%S')"
  fi

  mkdir -p "$DEST"
  rsync -a --delete "$SRC" "$DEST"

  # 注意：變數後緊接全形字時一律用 ${VAR} 大括號界定，否則舊 bash 3.2 在 set -u
  # 下會把後面那個多位元組字元的位元組誤讀進變數名，報 unbound variable。
  echo "已從 GitHub 拉回本機，備份於 ${BACKUP_DIR}，時間 $(date '+%Y-%m-%d %H:%M')"

  # 依賴自動同步：若這次 pull 有改到任何 skill 的 requirements.txt，就自動 pip 安裝，
  # 這樣平常更新只要跑本腳本一條。沒變動就跳過（多數 skill 更新不碰依賴）。
  if [ "$before_rev" != "$after_rev" ]; then
    # 只挑 .claude/skills/<skill>/requirements.txt 這種檔案的變動（新增或修改）
    local changed_reqs
    changed_reqs="$(git diff --name-only "$before_rev" "$after_rev" -- .claude/skills \
                    | grep -E '(^|/)requirements\.txt$' || true)"
    if [ -n "$changed_reqs" ]; then
      local PY
      PY="$(command -v python3 || true)"
      if [ -z "$PY" ]; then
        echo "⚠️  偵測到依賴有變動，但找不到 python3，請自行安裝後手動 pip install。"
      else
        echo "$changed_reqs" | while IFS= read -r rel; do
          [ -n "$rel" ] || continue
          dest_req="$DEST${rel#.claude/skills/}"   # repo 路徑 → 本機 ~/.claude/skills 對應路徑
          [ -f "$dest_req" ] || continue                 # 上游刪掉該檔就跳過
          echo "偵測到依賴變動，安裝：${dest_req}"
          "$PY" -m pip install -r "$dest_req" \
            || echo "⚠️  安裝失敗，請手動執行：$PY -m pip install -r \"$dest_req\""
        done
      fi
    fi
  fi
}

main "$@"
