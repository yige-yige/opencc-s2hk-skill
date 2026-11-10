# 安装指南

## 方式 A：仓库安装器（推荐）

适用于 macOS / Linux。安装器会处理 Python 依赖、分发 skill，并创建 `~/.local/bin/s2hk`。

```bash
git clone https://github.com/yige-yige/opencc-s2hk-skill.git
bash opencc-s2hk-skill/install.sh --list
bash opencc-s2hk-skill/install.sh
```

| 参数 | 作用 |
|---|---|
| `--list` | 只显示安装目标，不写文件 |
| `--all` | 同时创建未检测到的 agent 目录 |
| `--link` | 在专属目录中使用软链接而非复制 |
| `--no-deps` | 跳过 Python 依赖安装 |
| `--add-path` | 把 `~/.local/bin` 追加到 `~/.zshrc` |
| `--core-only` | 保留的 v1 兼容参数 |

安装器会备份已有的同名 skill 和命令。卸载时只删除能确认由本项目创建的命令与 XDG 链接，并恢复安装前的同名命令。

## 方式 B：skills CLI

```bash
npx skills add yige-yige/opencc-s2hk-skill --list
npx skills add yige-yige/opencc-s2hk-skill --all -g
```

skills CLI 只分发 skill 文件，不安装 Python 依赖。首次使用前执行：

```bash
python3 -m pip install opencc lxml python-pptx
```

如果系统 Python 受 PEP 668 限制，改用方式 A，让安装器创建独立 venv。

## 方式 C：手动复制

```bash
mkdir -p ~/.agents/skills
cp -R skills/opencc-s2hk ~/.agents/skills/
python3 -m pip install opencc lxml python-pptx
```

常见目录：

| 客户端 | 常见 skill 目录 |
|---|---|
| Agent Skills 兼容客户端 | `~/.agents/skills/` |
| Claude Code | `~/.claude/skills/` |
| CodeBuddy | `~/.codebuddy/skills/` |
| WorkBuddy | `~/.workbuddy/skills/` |

客户端可能支持项目级和全局级两种安装。以对应客户端的当前文档为准。

## Windows

捆绑的 Bash 安装器不支持 Windows。可手动复制 skill，然后执行：

```powershell
py -m pip install opencc lxml python-pptx
py <skill-dir>\scripts\s2hk_convert.py --text "用户启用多租户"
```

## 卸载

```bash
bash opencc-s2hk-skill/uninstall.sh
```

加 `--keep-env` 可保留独立 Python 环境。`~/.zshrc` 中由 `--add-path` 追加的 PATH 行需手动删除。
