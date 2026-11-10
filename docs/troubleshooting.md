# 排错指南

## `command not found: python3`

安装 Python 3 后重试。macOS 可使用 `brew install python`；Linux 使用系统包管理器。

## `externally-managed-environment`

这是 PEP 668 对系统 Python 的保护。推荐运行仓库 `install.sh`，让它创建独立 venv，不要手动破坏系统 Python。

## `ModuleNotFoundError: opencc`

用 skills CLI 或手动复制时，Python 依赖不会自动安装。可改用仓库安装器，或在当前 Python 环境中安装：

```bash
python3 -m pip install opencc lxml python-pptx
```

## Agent 没有识别 skill

1. 确认目录层级是 `<skills-dir>/opencc-s2hk/SKILL.md`，没有多套一层。
2. 完全退出并重新打开 agent，或新建会话。
3. 确认安装位置是当前客户端支持的项目或全局 skill 目录。
4. 如果客户端不支持 Agent Skills，把 `AGENT_SNIPPET.md` 复制到它的自定义指令中，并使用 `~/.local/bin/s2hk`。

常见安装目录见 [安装指南](installation.md)。

## PPTX 仍有少量简体

默认已覆盖正文、版式、母版、备注、SmartArt 和图表。常见剩余来源：

- 文字在图片或截图中。
- 文字在内嵌 OLE 对象中。
- 本次命令使用了缩小的 `--scope`。

可重跑并加 `--verify` 检查。

## 某些字看起来像简体

香港字形会合理保留 `户`、`峰`、`兌` 等字。不要用固定简体字表判定残留；使用 `--verify` 的二次转换检查。

## 「程序」没有变成「程式」

OpenCC 处理字形，不做港台词汇本地化。这类词汇替换需要另行编辑。

## 字体没有按预期显示

PPTX 中写入的是字体名。如果目标电脑未安装该字体，PowerPoint 会使用替代字体。请在打开文件的电脑上安装相同字体，或改用跨平台常见字体。
