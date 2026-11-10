#!/usr/bin/env bash
# 卸载简体 → 香港繁体 Skill：清理所有 agent 目录、全局命令与独立运行环境
set -uo pipefail

SKILL_NAMES=(opencc-s2hk pptx-s2t-font-convert)   # 后者为 v1 遗留，一并清理
CANON="$HOME/.agents/skills"
XDG_CANON="${XDG_CONFIG_HOME:-$HOME/.config}/agents/skills"
BIN_DIR="$HOME/.local/bin"
ENV_DIR="$HOME/.local/share/opencc-s2hk-skill"
BIN_BACKUP_DIR="$ENV_DIR/bin-backups"
XDG_LINK_MARKER="$ENV_DIR/managed-xdg-link"

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

# 只删除安装器记录且仍指向原目标的 XDG 软链接。
if [ -f "$XDG_LINK_MARKER" ]; then
  MANAGED_XDG="$(sed -n '1p' "$XDG_LINK_MARKER")"
  MANAGED_TARGET="$(sed -n '2p' "$XDG_LINK_MARKER")"
  if [ "$MANAGED_XDG" = "$XDG_CANON" ] && [ -L "$XDG_CANON" ] \
      && [ "$(readlink "$XDG_CANON")" = "$MANAGED_TARGET" ]; then
    rm -f "$XDG_CANON" && printf '  已删除本项目创建的软链接 %s\n' "$XDG_CANON"
  else
    printf '  已保留非本项目所有或已变更的 XDG 路径 %s\n' "$XDG_CANON"
  fi
  rm -f "$XDG_LINK_MARKER"
fi

is_managed_command() {
  _cmd="$1"
  [ -f "$_cmd" ] && [ ! -L "$_cmd" ] || return 1
  grep -q '^# managed-by: opencc-s2hk-skill$' "$_cmd" 2>/dev/null \
    || grep -q '^exec .*opencc-s2hk/scripts/s2hk_convert.py' "$_cmd" 2>/dev/null
}

for f in s2hk s2hk-pptx-font; do
  cmd="$BIN_DIR/$f"; backup="$BIN_BACKUP_DIR/$f"
  if [ -e "$cmd" ] || [ -L "$cmd" ]; then
    if is_managed_command "$cmd"; then
      rm -f "$cmd" && printf '  已删除本项目创建的命令 %s\n' "$cmd"
    else
      printf '  已保留非本项目所有的命令 %s\n' "$cmd"
    fi
  fi
  if [ ! -e "$cmd" ] && [ ! -L "$cmd" ] \
      && { [ -e "$backup" ] || [ -L "$backup" ]; }; then
    mv "$backup" "$cmd" && printf '  已恢复安装前的命令 %s\n' "$cmd"
  fi
done

for f in s2hk s2hk-pptx-font; do
  if [ -e "$BIN_BACKUP_DIR/$f" ] || [ -L "$BIN_BACKUP_DIR/$f" ]; then
    KEEP_ENV=1
    printf '  仍有同名命令备份未恢复，已保留运行环境：%s\n' "$BIN_BACKUP_DIR/$f"
  fi
done

if [ "$KEEP_ENV" -eq 0 ] && [ -d "$ENV_DIR" ]; then
  rm -rf "$ENV_DIR" && printf '  已删除独立运行环境 %s\n' "$ENV_DIR"
fi

printf '\n卸载完成。\n'
printf '备注：~/.local/share/opencc-s2hk-skill/backups/ 下的历史备份与 ~/.zshrc 里的 PATH 行（如有）需手动清理。\n'
