#!/usr/bin/env bash
# 卸载简体 → 香港繁体 Skill：清理所有 agent 目录、全局命令与独立运行环境
set -uo pipefail

SKILL_NAMES=(opencc-s2hk pptx-s2t-font-convert)   # 后者为 v1 遗留，一并清理
CANON="$HOME/.agents/skills"
XDG_CANON="${XDG_CONFIG_HOME:-$HOME/.config}/agents/skills"
BIN_DIR="$HOME/.local/bin"
ENV_DIR="$HOME/.local/share/opencc-s2hk-skill"

DIRS=(
  "$CANON"
  "$HOME/.codebuddy/skills"
  "$HOME/.workbuddy/skills"
  "$HOME/.claude/skills"
  "$HOME/.cursor/skills"
  "$HOME/.gemini/skills"
  "$HOME/.copilot/skills"
  "$HOME/.config/opencode/skills"
  "$HOME/.config/goose/skills"
  "$HOME/.codeium/windsurf/skills"
  "$HOME/.roo/skills"
  "$HOME/.kilo/skills"
  "$HOME/.kiro/skills"
  "$HOME/.factory/skills"
  "$HOME/.qwen/skills"
  "$HOME/.iflow/skills"
  "$HOME/.lingma/skills"
  "$HOME/.trae/skills"
  "$HOME/.trae-cn/skills"
  "$HOME/.qoder/skills"
  "$HOME/.qoder-cn/skills"
  "$HOME/.junie/skills"
  "$HOME/.openclaw/skills"
)

KEEP_ENV=0
[ "${1:-}" = "--keep-env" ] && KEEP_ENV=1

for dir in "${DIRS[@]}"; do
  for n in "${SKILL_NAMES[@]}"; do
    if [ -e "$dir/$n" ] || [ -L "$dir/$n" ]; then
      rm -rf "$dir/$n" && printf '  已删除 %s\n' "$dir/$n"
    fi
  done
  # 旧版安装器可能把备份留在 skills 目录里，一并清掉（它们会造成技能重复）
  for d in "$dir"/opencc-s2hk.bak.* "$dir"/pptx-s2t-font-convert.bak.*; do
    [ -e "$d" ] || continue
    rm -rf "$d" && printf '  已删除遗留备份 %s\n' "$d"
  done
done

# XDG 软链接（指向 CANON 时才删）
if [ -L "$XDG_CANON" ]; then
  rm -f "$XDG_CANON" && printf '  已删除软链接 %s\n' "$XDG_CANON"
fi

for f in s2hk s2hk-pptx-font; do
  if [ -e "$BIN_DIR/$f" ]; then
    rm -f "$BIN_DIR/$f" && printf '  已删除命令 %s\n' "$BIN_DIR/$f"
  fi
done

if [ "$KEEP_ENV" -eq 0 ] && [ -d "$ENV_DIR" ]; then
  rm -rf "$ENV_DIR" && printf '  已删除独立运行环境 %s\n' "$ENV_DIR"
fi

printf '\n卸载完成。\n'
printf '备注：~/.local/share/opencc-s2hk-skill/backups/ 下的历史备份与 ~/.zshrc 里的 PATH 行（如有）需手动清理。\n'
