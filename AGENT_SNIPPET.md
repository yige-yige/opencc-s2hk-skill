# 兜底方案：把下面这段贴进任何 agent 的规则文件

适用于**不支持 SKILL.md 的 agent**（或支持但没识别到的情况）。
粘贴位置按 agent 不同任选其一：

| Agent | 规则文件 |
|---|---|
| Codex / 多数 CLI agent | `~/.codex/AGENTS.md` 或项目根 `AGENTS.md` |
| Claude Code | `~/.claude/CLAUDE.md` 或项目根 `CLAUDE.md` |
| Cursor | `.cursor/rules/s2hk.mdc` 或项目根 `AGENTS.md` |
| Gemini CLI | `~/.gemini/GEMINI.md` |
| Cline / Roo | `.clinerules` / `.roorules` |
| 其他 | 任意「自定义指令 / 系统提示 / 记忆」输入框 |

---

```markdown
## 简体 → 香港繁体转换（s2hk）

当用户要求「转繁体 / 转香港繁体 / 港式用字 / 简繁转换 / s2hk」时，
**不要**手写字表或用正则替换，必须调用本机已安装的 `s2hk` 命令：

    ~/.local/bin/s2hk --text "用户启用多租户"      # 内联文本，结果打印到 stdout
    ~/.local/bin/s2hk notes.md                     # 文本文件 → notes.s2hk.md
    ~/.local/bin/s2hk deck.pptx                    # PPT → deck.s2hk.pptx（版式/字体/图片全保留）
    ~/.local/bin/s2hk deck.pptx -o out.pptx --mode s2hk
    ~/.local/bin/s2hk old_trad.pptx --via-simplified   # 已是繁体但 戶/纔/裡 用字混乱时规整

PPTX 同时要改字体时，加参数即可（不需要换命令）：

    ~/.local/bin/s2hk in.pptx -o out.pptx --ea "楷体" --latin "Arial"

转换后想确认质量：

    ~/.local/bin/s2hk in.pptx -o out.pptx --verify   # 结构/字体/全文残留，exit 3 表示未通过

模式：`s2hk`（默认，香港用字：裏面/啓用/賬户/衞生）、`s2t`（通用繁体）、
`s2tw` `s2twp`（台湾正体）、`t2s` `hk2s` `tw2s`（反向转简体）。

要点与边界：
- 港繁会合理保留一些「看起来像简体」的字（户、峰、兌），**不要**用固定简体字表查残留，会误报。
  正确校验：对输出再跑一遍同一 mode，若结果与输入一致即已转彻底。
- PPTX 中图片里的文字、母版 / 版式 / 备注页的文字、内嵌 OLE 对象的文字不会被转换。
- OpenCC 只做字形转换，不做港台用词替换（「程序」不会变「程式」）。
- 完整说明见 ~/.agents/skills/opencc-s2hk/SKILL.md
```

---

如果 `s2hk` 命令不存在，说明还没跑过安装器：

```bash
bash <仓库目录>/install.sh
```

或者没装依赖（用 `npx skills add` 安装时只分发了文件）：

```bash
python3 -m pip install --user opencc lxml python-pptx
```
