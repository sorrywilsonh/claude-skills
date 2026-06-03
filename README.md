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

## 在新電腦還原到全域（讓本機 CLI 全專案都能用）

```bash
git clone <this-repo> ~/projects/claude-skills
cp -R ~/projects/claude-skills/.claude/skills/. ~/.claude/skills/
```

## 在 web 上使用

在 claude.ai/code 開啟需要這些 skill 的 repo，並確保該 repo 的 `.claude/skills/` 內含這些 skill（直接複製或用 submodule）。
