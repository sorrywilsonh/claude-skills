# claude-skills

我的 Claude Code skills 備份與共用 repo。

## 用途

- **雲端備份**：保存個人 skills，避免本機誤刪或重灌遺失。
- **跨平台共用**：skills 放在 `.claude/skills/`，可被本機 Claude Code CLI 與 Claude Code on web（claude.ai/code）讀取。

> 注意：claude.ai 純聊天網站目前不支援這些 Claude Code skills。

## 結構

```
.claude/skills/
└── <skill-name>/SKILL.md   # 每個 skill 一個資料夾
```

## 外部來源 Skill

| Skill | 來源 | 說明 |
|---|---|---|
| `grill-me` | [mattpocock/skills](https://github.com/mattpocock/skills/tree/3cca18b368ae95cdbdebbff572ccafa662551015/skills/productivity/grill-me)(MIT) | 手動觸發(`disable-model-invocation`),呼叫 `grilling` 執行實際邏輯 |
| `grilling` | [mattpocock/skills](https://github.com/mattpocock/skills/tree/3cca18b368ae95cdbdebbff572ccafa662551015/skills/productivity/grilling)(MIT) | `grill-me` 依賴的實際 skill:針對計畫/決策連續追問,逐輪列出問題與建議答案 |
| `brainstorming`、`dispatching-parallel-agents`、`executing-plans`、`finishing-a-development-branch`、`receiving-code-review`、`requesting-code-review`、`subagent-driven-development`、`systematic-debugging`、`test-driven-development`、`using-git-worktrees`、`using-superpowers`、`verification-before-completion`、`writing-plans`、`writing-skills` | [obra/superpowers](https://github.com/obra/superpowers)(MIT,Jesse Vincent) | 語言無關的工程紀律 skill 包。這裡存的是早期匯入時的快照,上游持續在更新,沒有自動同步 |
| `ppt-master` | [hugohe3/ppt-master](https://github.com/hugohe3/ppt-master)(MIT,Hugo He) | AI 簡報產生工具。已同步到上游 **v6.4.0**(commit `b44cdfa`):路由式工作流(Generate / Create Template / Edit Native PPTX),Brand/Style/Layout/Deck workspace、原生可編輯 PPTX、19 種視覺風格、旁白/影片。上游 skill 現位於其 repo 的 `skills/ppt-master/`,已對應到本 repo 的 `.claude/skills/ppt-master/`;沒有自動同步。SKILL.md 正文頂部另加了一段本機 LOCAL SETUP 說明(不影響上游的完整性校驗)。 |

`grill-me` 只是個轉發用的別名,兩個資料夾要一起存在才能用。以上每個外部 skill 的資料夾裡都放了來源的 MIT LICENSE。

## 在新電腦還原到全域（讓本機 CLI 全專案都能用）

```bash
git clone <this-repo> ~/projects/claude-skills
cp -R ~/projects/claude-skills/.claude/skills/. ~/.claude/skills/
```

## 雙向同步腳本

repo 假設本機 clone 在 `~/projects/claude-skills`。

| 腳本 | 方向 | 用途 |
|---|---|---|
| `sync.sh` | 本機 `~/.claude/skills/` → repo → GitHub | 把本機修改推上雲端(鏡像同步，含刪除) |
| `pull.sh` | GitHub → repo → 本機 `~/.claude/skills/` | 把雲端最新版拉回本機(鏡像同步，含刪除) |

`pull.sh` 執行前會先確認 repo 沒有未推送的變動(避免弄丟),覆蓋本機前也會先備份一份到
`~/.claude/skills-backups/`，以防尚未跑過 `sync.sh` 的本機修改被覆蓋掉。

多台電腦協作時的建議流程：**改完 skill → 跑 `sync.sh` 推上去 → 換到別台電腦先跑 `pull.sh` 拉下來，再繼續改**。同一時間只在一台電腦上改動,避免雙邊都有未同步的修改互相覆蓋。

## 在 web 上使用

在 claude.ai/code 開啟需要這些 skill 的 repo，並確保該 repo 的 `.claude/skills/` 內含這些 skill（直接複製或用 submodule）。
