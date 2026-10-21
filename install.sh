#!/usr/bin/env bash
# =============================================================================
#  简体 → 香港繁体 Skill · 通用安装器（macOS / Linux）
#
#  三层覆盖，确保任何 agent 都能用：
#   ① 标准共享目录 ~/.agents/skills  → Codex / Cursor / Gemini CLI / Copilot /
#                                      Cline / Zed / Warp / Amp / OpenCode ...
#   ② 各家专属目录                   → CodeBuddy / WorkBuddy / Claude Code ...
#   ③ 全局命令 ~/.local/bin/s2hk     → 兜底：只要 agent 能跑终端命令就能用
#
#  用法：
#    bash install.sh              # 安装（推荐）
#    bash install.sh --list       # 只检测，不写任何文件
#    bash install.sh --all        # 连未检测到的 agent 目录也一起建
#    bash install.sh --link       # 用软链接代替复制（升级只需改一处）
#    bash install.sh --core-only  # 只装核心 skill，不装 PPT 字体增强版
#    bash install.sh --no-deps    # 跳过 Python 依赖安装
#    bash install.sh --add-path   # 自动把 ~/.local/bin 写进 ~/.zshrc
# =============================================================================
set -uo pipefail

PACK_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SKILLS_SRC="$PACK_DIR/skills"
CANON="$HOME/.agents/skills"                              # Agent Skills 标准位置
XDG_CANON="${XDG_CONFIG_HOME:-$HOME/.config}/agents/skills"
VENV_DIR="$HOME/.local/share/opencc-s2hk-skill/venv"
BIN_DIR="$HOME/.local/bin"
PKGS=(opencc lxml python-pptx)

MODE="install"; LINK=0; FORCE_ALL=0; CORE_ONLY=0; WITH_DEPS=1; ADD_PATH=0

while [ $# -gt 0 ]; do
  case "$1" in
    --list) MODE="list"; shift ;;
    --link) LINK=1; shift ;;
    --all) FORCE_ALL=1; shift ;;
    --core-only) CORE_ONLY=1; shift ;;
    --no-deps) WITH_DEPS=0; shift ;;
    --add-path) ADD_PATH=1; shift ;;
    -h|--help) sed -n '2,22p' "$0"; exit 0 ;;
    *) echo "未知参数：$1（--help 查看用法）" >&2; exit 2 ;;
  esac
done

say()  { printf '%s\n' "$*"; }
ok()   { printf '  [ok]   %s\n' "$*"; }
warn() { printf '  [warn] %s\n' "$*"; }
die()  { printf '\n[失败] %s\n' "$*" >&2; exit 1; }

# 检测路径|技能目录|显示名|是否无条件安装
AGENT_MAP=(
  "$HOME/.codebuddy|$HOME/.codebuddy/skills|CodeBuddy|1"
  "$HOME/.workbuddy|$HOME/.workbuddy/skills|WorkBuddy|1"
  "$HOME/.claude|$HOME/.claude/skills|Claude Code|0"
  "$HOME/.cursor|$HOME/.cursor/skills|Cursor|0"
  "$HOME/.gemini|$HOME/.gemini/skills|Gemini CLI|0"
  "$HOME/.copilot|$HOME/.copilot/skills|GitHub Copilot|0"
  "$HOME/.config/opencode|$HOME/.config/opencode/skills|OpenCode|0"
  "$HOME/.config/goose|$HOME/.config/goose/skills|Goose|0"
  "$HOME/.codeium|$HOME/.codeium/windsurf/skills|Windsurf|0"
  "$HOME/.roo|$HOME/.roo/skills|Roo Code|0"
  "$HOME/.kilo|$HOME/.kilo/skills|Kilo Code|0"
  "$HOME/.kiro|$HOME/.kiro/skills|Kiro|0"
  "$HOME/.factory|$HOME/.factory/skills|Droid (Factory)|0"
  "$HOME/.qwen|$HOME/.qwen/skills|Qwen Code|0"
  "$HOME/.iflow|$HOME/.iflow/skills|iFlow CLI|0"
  "$HOME/.lingma|$HOME/.lingma/skills|通义灵码|0"
  "$HOME/.trae|$HOME/.trae/skills|Trae|0"
  "$HOME/.trae-cn|$HOME/.trae-cn/skills|Trae CN|0"
  "$HOME/.qoder|$HOME/.qoder/skills|Qoder|0"
  "$HOME/.qoder-cn|$HOME/.qoder-cn/skills|Qoder CN|0"
  "$HOME/.junie|$HOME/.junie/skills|Junie|0"
  "$HOME/.openclaw|$HOME/.openclaw/skills|OpenClaw|0"
)

SKILL_NAMES=(opencc-s2hk)
[ "$CORE_ONLY" -eq 0 ] && SKILL_NAMES+=(pptx-s2t-font-convert)
for n in "${SKILL_NAMES[@]}"; do
  [ -f "$SKILLS_SRC/$n/SKILL.md" ] || die "安装包不完整，缺少 skills/$n/SKILL.md"
done

# ------------------------------------------------------- 解析本机要安装的位置
TARGETS=(); LABELS=()
for row in "${AGENT_MAP[@]}"; do
  det="${row%%|*}";        rest="${row#*|}"
  dir="${rest%%|*}";       rest="${rest#*|}"
  label="${rest%%|*}";     always="${rest##*|}"
  if [ "$always" = "1" ] || [ "$FORCE_ALL" -eq 1 ] || [ -d "$det" ]; then
    TARGETS+=("$dir"); LABELS+=("$label")
  fi
done

say "================================================"
say " 简体 → 香港繁体 Skill · 通用安装器"
say "================================================"
say ""
say "将要安装的 skill：${SKILL_NAMES[*]}"
say ""
say "① 标准共享目录（Codex / Cursor / Gemini CLI / Copilot / Cline / Zed / Warp 等都读这里）"
say "   $CANON"
say ""
say "② 专属目录："
i=0
while [ $i -lt ${#TARGETS[@]} ]; do
  printf '   %-18s %s\n' "${LABELS[$i]}" "${TARGETS[$i]}"
  i=$((i + 1))
done
say ""
say "③ 全局命令：$BIN_DIR/s2hk"

if [ "$MODE" = "list" ]; then
  say ""
  say "（--list 模式：未写入任何文件）"
  exit 0
fi

# ------------------------------------------------------------ 1. Python 依赖
PYBIN=""
say ""
say "[1/4] Python 依赖（opencc / lxml / python-pptx）"
if [ "$WITH_DEPS" -eq 1 ]; then
  SYS_PY="$(command -v python3 2>/dev/null)"
  [ -n "$SYS_PY" ] || die "未找到 python3。先执行  xcode-select --install  （或 brew install python）再重跑本脚本"
  ok "python3 → $SYS_PY ($("$SYS_PY" -V 2>&1))"

  deps_ok() { "$1" -c "import opencc, lxml, pptx" >/dev/null 2>&1; }

  if deps_ok "$SYS_PY"; then
    PYBIN="$SYS_PY"; ok "依赖已齐全"
  else
    # 注意：数组不能为空，macOS 自带 bash 3.2 在 set -u 下会把空数组报成 unbound
    for mode in plain user break; do
      case "$mode" in
        plain) ARGS=("${PKGS[@]}") ;;
        user)  ARGS=(--user "${PKGS[@]}") ;;
        break) ARGS=(--user --break-system-packages "${PKGS[@]}") ;;
      esac
      "$SYS_PY" -m pip install --quiet --disable-pip-version-check "${ARGS[@]}" >/dev/null 2>&1
      if deps_ok "$SYS_PY"; then PYBIN="$SYS_PY"; ok "已装入系统 python3（模式：$mode）"; break; fi
    done

    if [ -z "$PYBIN" ]; then
      warn "系统 pip 受限（PEP 668），改建独立运行环境"
      rm -rf "$VENV_DIR"; mkdir -p "$(dirname "$VENV_DIR")"
      "$SYS_PY" -m venv "$VENV_DIR" >/dev/null 2>&1 || die "创建 venv 失败"
      "$VENV_DIR/bin/python" -m pip install --quiet --upgrade pip >/dev/null 2>&1
      "$VENV_DIR/bin/python" -m pip install --quiet "${PKGS[@]}" >/dev/null 2>&1
      deps_ok "$VENV_DIR/bin/python" && { PYBIN="$VENV_DIR/bin/python"; ok "独立环境就绪 → $VENV_DIR"; }
    fi

    if [ -z "$PYBIN" ]; then   # opencc 无预编译包时换纯 Python 实现
      CAND="$VENV_DIR/bin/python"; [ -x "$CAND" ] || CAND="$SYS_PY"
      "$CAND" -m pip install --quiet opencc-python-reimplemented >/dev/null 2>&1
      deps_ok "$CAND" && { PYBIN="$CAND"; ok "已改用 opencc-python-reimplemented"; }
    fi

    [ -n "$PYBIN" ] || die "依赖安装失败。请手动执行：$SYS_PY -m pip install --user ${PKGS[*]}"
  fi
else
  warn "跳过依赖安装（--no-deps）"
fi

# -------------------------------------------- 2. 装到标准共享目录（真实文件）
say ""
say "[2/4] 安装到标准共享目录"
mkdir -p "$CANON" || die "无法创建 $CANON"
for n in "${SKILL_NAMES[@]}"; do
  if [ -e "$CANON/$n" ] || [ -L "$CANON/$n" ]; then
    BAK="$CANON/$n.bak.$(date +%Y%m%d%H%M%S)"
    mv "$CANON/$n" "$BAK" && warn "$n 已存在，旧版备份到 $(basename "$BAK")"
  fi
  cp -R "$SKILLS_SRC/$n" "$CANON/" || die "拷贝 $n 失败"
  [ -n "$PYBIN" ] && printf '%s\n' "$PYBIN" > "$CANON/$n/scripts/.python"
  ok "$n → $CANON/$n"
done
# 少数 agent（Amp / Replit 等）读 XDG 路径
if [ "$XDG_CANON" != "$CANON" ]; then
  mkdir -p "$(dirname "$XDG_CANON")" 2>/dev/null
  if [ ! -e "$XDG_CANON" ]; then
    ln -s "$CANON" "$XDG_CANON" 2>/dev/null && ok "XDG 路径已链接 → $XDG_CANON"
  fi
fi

# ------------------------------------------------- 3. 分发到各家专属目录
say ""
if [ "$LINK" -eq 1 ]; then say "[3/4] 分发到各家专属目录（软链接）"; else say "[3/4] 分发到各家专属目录（复制）"; fi
INSTALLED=()
i=0
while [ $i -lt ${#TARGETS[@]} ]; do
  dir="${TARGETS[$i]}"; label="${LABELS[$i]}"; i=$((i + 1))
  [ "$dir" = "$CANON" ] && continue
  mkdir -p "$dir" 2>/dev/null || { warn "$label：无法创建 $dir，跳过"; continue; }
  for n in "${SKILL_NAMES[@]}"; do
    if [ -e "$dir/$n" ] || [ -L "$dir/$n" ]; then
      BAK="$dir/$n.bak.$(date +%Y%m%d%H%M%S)"
      mv "$dir/$n" "$BAK" >/dev/null 2>&1 && warn "$label：旧版备份到 $(basename "$BAK")"
    fi
    if [ "$LINK" -eq 1 ]; then
      ln -s "$CANON/$n" "$dir/$n" || { warn "$label：链接 $n 失败"; continue; }
    else
      cp -R "$CANON/$n" "$dir/" || { warn "$label：拷贝 $n 失败"; continue; }
    fi
  done
  ok "$label → $dir"
  INSTALLED+=("$label  ($dir)")
done

# ----------------------------------------------------- 4. 全局命令 + 自检
say ""
say "[4/4] 全局命令与自检"
mkdir -p "$BIN_DIR" || die "无法创建 $BIN_DIR"
RUN_PY="${PYBIN:-python3}"

cat > "$BIN_DIR/s2hk" <<EOF
#!/bin/sh
# 简体 → 香港繁体（文本 / Markdown / PPTX）。任何能执行终端命令的 agent 都可调用。
exec "$RUN_PY" "$CANON/opencc-s2hk/scripts/s2hk_convert.py" "\$@"
EOF
chmod +x "$BIN_DIR/s2hk"
ok "s2hk → $BIN_DIR/s2hk"

if [ "$CORE_ONLY" -eq 0 ]; then
  cat > "$BIN_DIR/s2hk-pptx-font" <<EOF
#!/bin/sh
# PPTX 繁简转换 + 中英文字体统一。用法：s2hk-pptx-font in.pptx out.pptx --ea 楷体 --latin Arial --mode s2hk
exec "$RUN_PY" "$CANON/pptx-s2t-font-convert/scripts/convert.py" "\$@"
EOF
  chmod +x "$BIN_DIR/s2hk-pptx-font"
  ok "s2hk-pptx-font → $BIN_DIR/s2hk-pptx-font"
fi

case ":$PATH:" in
  *":$BIN_DIR:"*) ok "$BIN_DIR 已在 PATH 中" ;;
  *)
    if [ "$ADD_PATH" -eq 1 ]; then
      printf '\n# added by opencc-s2hk skill installer\nexport PATH="%s:$PATH"\n' "$BIN_DIR" >> "$HOME/.zshrc"
      ok "已把 $BIN_DIR 写入 ~/.zshrc（新开终端生效）"
    else
      warn "$BIN_DIR 不在 PATH 中。可重跑本脚本加 --add-path，或直接用绝对路径 $BIN_DIR/s2hk"
    fi ;;
esac

if [ -x "$BIN_DIR/s2hk" ]; then
  if RESULT="$("$BIN_DIR/s2hk" --text "用户启用多租户数据隔离机制" 2>&1)"; then
    ok "自检通过：用户启用多租户数据隔离机制 → $RESULT"
  else
    warn "自检未通过，输出如下："
    printf '%s\n' "$RESULT"
    die "请把以上报错截图发给分享者"
  fi
fi

say ""
say "================================================"
say " 安装完成"
say "================================================"
say "共享目录（覆盖 Codex / Cursor / Gemini CLI / Copilot / Cline / Zed / Warp / Amp / OpenCode 等）："
say "  $CANON"
say "专属目录："
for d in "${INSTALLED[@]}"; do say "  - $d"; done
say "全局命令：$BIN_DIR/s2hk   （示例：s2hk deck.pptx）"
say ""
say "下一步：重启 / 新开 agent 会话，然后直接说「把这个 PPT 转成香港繁体」。"
say "若某个 agent 不认 skill，把 $PACK_DIR/AGENT_SNIPPET.md 的内容贴进它的规则文件即可。"
say ""
say "卸载：bash \"$PACK_DIR/uninstall.sh\""
