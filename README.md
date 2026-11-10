# opencc-s2hk · 简体 → 香港繁体 Agent Skill

把文本、Markdown 和 PPTX 中的简体中文转换为**香港繁体（OpenCC `s2hk`）**。PPTX 保留版式、配色、图片与表格，还可选统一中英文字体。

这是一个符合 [Agent Skills 规范](https://agentskills.io/specification) 的自包含 skill；也可在任何能执行终端命令的 agent 中通过 `s2hk` 使用。

## 支持范围

| 输入 | 支持 | 说明 |
|---|---|---|
| 内联文本 | ✅ | 直接返回转换结果 |
| `.txt` `.md` `.csv` `.srt` | ✅ | UTF-8 纯文本 |
| `.pptx` | ✅ | 正文、版式、母版、备注、SmartArt、图表与文件属性 |
| `.docx` `.xlsx` `.pdf` `.key` | ❌ | 需改用相应的专用工具 |

捆绑安装器支持 macOS / Linux；Windows 可手动安装 Python 依赖后直接调用脚本。

## 快速安装

推荐使用仓库安装器：它会处理 Python 依赖、分发 skill，并创建 `s2hk` 命令。

```bash
git clone https://github.com/yige-yige/opencc-s2hk-skill.git
bash opencc-s2hk-skill/install.sh --list
bash opencc-s2hk-skill/install.sh
```

安装完成后，重启 agent 或新开会话。使用 skills CLI、手动复制、Windows 安装与卸载方法见 [安装指南](docs/installation.md)。

## 怎么用

在 agent 里用自然语言提出需求即可：

| 你说 | 结果 |
|---|---|
| 「把这个 PPT 转成香港繁体」 | 输出 `xxx.s2hk.pptx`，保留版式与媒体 |
| 「这份 Markdown 转香港繁体」 | 输出 `xxx.s2hk.md` |
| 「这段话转成港式用字」 | 直接返回转换后的文本 |
| 「PPT 转繁体，中文用楷体、英文用 Arial」 | 转换文字并统一字体 |
| 「这份繁体文档用字混乱，统一成香港写法」 | 经简体中转后规整为港式字形 |

也可直接使用命令行：

```bash
s2hk --text "用户启用多租户"
s2hk notes.md
s2hk deck.pptx
s2hk deck.pptx -o out.pptx --ea "楷体" --latin "Arial"
s2hk old_trad.pptx --via-simplified
s2hk deck.pptx -o out.pptx --verify
```

完整的模式、参数、PPTX 处理范围与质检说明见 [CLI 手册](docs/cli.md)。

## 已知边界

- 图片、截图、视频封面中的文字不会被修改。
- 内嵌 OLE 对象（如嵌入的 Excel）内部文字不在处理范围。
- OpenCC 只转换字形，不会把「程序」自动本地化为「程式」。
- 默认不改字体；只有明确传入 `--ea` / `--latin` 时才会改字体。
- 输入与输出不能是同一文件；脚本会拒绝同路径写入。

## 文档

- [安装指南](docs/installation.md) —— skills CLI、仓库安装器、手动安装与 Agent 路径
- [CLI 手册](docs/cli.md) —— 转换模式、PPTX scope、字体、质检与退出码
- [实现与测试](docs/implementation.md) —— 处理原理、安全措施、仓库结构与回归测试
- [排错指南](docs/troubleshooting.md) —— 依赖、skill 识别、残留文字与常见问题
- [变更记录](CHANGELOG.md)

## 许可

基于 [OpenCC](https://github.com/BYVoid/OpenCC)、`lxml` 与 `python-pptx`。本仓库代码采用 [MIT License](LICENSE)。
