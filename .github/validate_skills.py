#!/usr/bin/env python3
"""仓库自检：校验 SKILL.md 规范 + 脚本可编译 + 转换结果可用。

在 CI 与本地均可直接运行：python3 .github/validate_skills.py
"""
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SKILLS_DIR = os.path.join(ROOT, "skills")
NAME_RE = re.compile(r"^[a-z0-9]+(-[a-z0-9]+)*$")

errors = []
checked = 0


def fail(msg):
    errors.append(msg)


def frontmatter(path):
    """返回 frontmatter 文本块；不做完整 YAML 解析，避免引入依赖。"""
    with open(path, encoding="utf-8") as f:
        lines = f.read().split("\n")
    if not lines or lines[0].strip() != "---":
        return None
    for i in range(1, len(lines)):
        if lines[i].strip() == "---":
            return "\n".join(lines[1:i])
    return None


def top_level_value(block, key):
    """取 frontmatter 顶层字段的值（支持跨行的折叠字符串）。"""
    m = re.search(r"^%s:(.*)$" % re.escape(key), block, re.MULTILINE)
    if not m:
        return None
    value = [m.group(1).strip()]
    rest = block[m.end():].split("\n")[1:]
    for line in rest:
        if re.match(r"^[A-Za-z0-9_-]+:", line):      # 下一个顶层字段
            break
        if line.strip():
            value.append(line.strip())
    return " ".join(value).strip().strip('"').strip()


# ---------------------------------------------------------------- SKILL.md
if not os.path.isdir(SKILLS_DIR):
    fail("缺少 skills/ 目录")

for name in sorted(os.listdir(SKILLS_DIR) if os.path.isdir(SKILLS_DIR) else []):
    skill_dir = os.path.join(SKILLS_DIR, name)
    if not os.path.isdir(skill_dir) or name.startswith("."):
        continue
    checked += 1
    md = os.path.join(skill_dir, "SKILL.md")
    if not os.path.isfile(md):
        fail("%s: 缺少 SKILL.md" % name)
        continue

    block = frontmatter(md)
    if block is None:
        fail("%s/SKILL.md: 第一行必须是 --- 且 frontmatter 要正确闭合" % name)
        continue

    declared = top_level_value(block, "name")
    if not declared:
        fail("%s/SKILL.md: frontmatter 缺少必填字段 name" % name)
    elif declared != name:
        fail("%s/SKILL.md: name=%r 与目录名不一致（部分平台要求严格相等）"
             % (name, declared))
    elif not NAME_RE.match(declared):
        fail("%s/SKILL.md: name 必须是小写 kebab-case" % name)

    desc = top_level_value(block, "description")
    if not desc:
        fail("%s/SKILL.md: frontmatter 缺少必填字段 description" % name)
    elif len(desc) < 40:
        fail("%s/SKILL.md: description 过短（%d 字符），隐式触发会不准"
             % (name, len(desc)))

    # 脚本可编译（只做语法检查，不落盘字节码）
    scripts = os.path.join(skill_dir, "scripts")
    if os.path.isdir(scripts):
        for fn in sorted(os.listdir(scripts)):
            if fn.endswith(".py"):
                p = os.path.join(scripts, fn)
                try:
                    with open(p, encoding="utf-8") as f:
                        compile(f.read(), p, "exec")
                except SyntaxError as e:
                    fail("%s/scripts/%s: 语法错误 %s" % (name, fn, e))

    # 安装器产物不应进仓库
    if os.path.isfile(os.path.join(scripts, ".python")):
        fail("%s/scripts/.python 不应提交（见 .gitignore）" % name)

if checked == 0:
    fail("skills/ 下没有发现任何 skill")

# ------------------------------------------------------------------ 转换自检
try:
    from opencc import OpenCC
except ImportError:
    print("[skip] 未安装 opencc，跳过转换自检")
else:
    cc = OpenCC("s2hk")
    src = "用户启用多租户数据隔离机制"
    out = cc.convert(src)
    if out == src:
        fail("转换自检失败：s2hk 没有产生任何变化")
    elif cc.convert(out) != out:
        fail("转换自检失败：输出不是幂等的（二次转换仍有变化）→ %r" % out)
    else:
        print("[ok] s2hk: %s → %s" % (src, out))

# ---------------------------------------------------------------------- 汇总
print("\n已检查 %d 个 skill" % checked)
if errors:
    print("\n发现 %d 个问题：" % len(errors))
    for e in errors:
        print("  - %s" % e)
    sys.exit(1)
print("全部通过")
