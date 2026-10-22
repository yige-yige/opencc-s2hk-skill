#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
s2hk_convert.py — Simplified→Traditional Chinese converter (Hong Kong variant by default),
built on OpenCC. Self-contained, portable, no external project structure required.

Handles three input kinds:
  1. Inline text     : --text "简体文字"
  2. Plain text file : .txt / .md / .csv / .srt / any UTF-8 text  → writes a converted copy
  3. PPTX file       : .pptx → converts all slide text, PRESERVING layout/fonts/images/tables
                       (only <a:t> text nodes in ppt/slides/slideN.xml are touched;
                        fonts, colors, shapes, OLE objects are left untouched)

Conversion modes (OpenCC):
  s2hk  (default) Simplified → Hong Kong Traditional   e.g. 裏面 啓用 覆盤 賬户 (keeps 户, not 戶)
  s2t             Simplified → Traditional (generic/OpenCC standard)
  s2tw            Simplified → Taiwan Traditional
  s2twp           Simplified → Taiwan Traditional + Taiwan phrase substitution
  t2s / hk2s / tw2s / tw2sp   reverse (→ Simplified)

Extra:
  --via-simplified   Normalize an ALREADY-Traditional (wrong/mixed variant) input back to
                     Simplified first (t2s), then convert to the target mode. Use this to
                     regularize 戶→户, 纔→才, 裡→裏, 啟→啓 into consistent HK forms.

Examples:
  python3 s2hk_convert.py --text "用户启用多租户数据隔离"
  python3 s2hk_convert.py notes.md                       # -> notes.s2hk.md
  python3 s2hk_convert.py notes.md -o notes_hk.md
  python3 s2hk_convert.py deck.pptx                       # -> deck.s2hk.pptx
  python3 s2hk_convert.py deck.pptx -o deck_hk.pptx --mode s2hk
  python3 s2hk_convert.py already_trad.pptx --mode s2hk --via-simplified

Dependencies:
  pip install opencc            # required for text conversion
  pip install lxml              # required ONLY for .pptx input
Verify install:
  python3 -c "import opencc; print('opencc ok')"

Residue verification tip (variant-aware — do NOT use a fixed simplified-char list):
  Re-run the SAME mode over the output; if convert(text) == text for all text, it's complete.
  HK legitimately keeps some 'simplified-looking' chars (户, 峰, 兌); a fixed dictionary
  gives false positives, so always use this re-convert method.
"""
import os
import sys
import argparse


def _bootstrap(mods):
    """Re-exec with the interpreter recorded by install.sh if deps are missing.

    install.sh writes the absolute path of a working interpreter (often a
    dedicated venv, needed when system pip is locked down by PEP 668
    'externally-managed-environment') into `<scripts>/.python`.
    """
    try:
        import importlib.util
        if all(importlib.util.find_spec(m) is not None for m in mods):
            return
    except Exception:
        pass
    if os.environ.get("SKILL_PY_REEXEC"):
        return
    candidates = []
    pin = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".python")
    if os.path.isfile(pin):
        try:
            with open(pin, encoding="utf-8") as f:
                candidates.append(f.read().strip())
        except OSError:
            pass
    # 通过 `npx skills` 等方式安装时没有 .python，回退到安装器的默认环境
    candidates.append(os.path.expanduser(
        "~/.local/share/opencc-s2hk-skill/venv/bin/python"))
    for py in candidates:
        if py and os.path.exists(py) and \
                os.path.realpath(py) != os.path.realpath(sys.executable):
            os.environ["SKILL_PY_REEXEC"] = "1"
            os.execv(py, [py, os.path.abspath(__file__)] + sys.argv[1:])


_bootstrap(("opencc", "lxml"))

# PPTX 处理核心与本脚本同目录（skill 自包含，不跨目录 import）
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pptx_core import (  # noqa: E402
    A, q, SCOPES, SCOPE_ORDER, parse_scopes, in_scope,
    safe_members, convert_paragraph_aware, convert_props, repack, verify,
)


def build_converter(mode, via_simplified):
    from opencc import OpenCC
    cc = OpenCC(mode)
    pre = OpenCC("t2s") if via_simplified else None

    def convert(s):
        if s is None:
            return s
        if pre is not None:
            s = pre.convert(s)
        return cc.convert(s)

    return convert


def convert_inline_text(text, convert):
    sys.stdout.write(convert(text) + "\n")


BINARY_HINTS = {
    ".docx": "Word", ".doc": "Word", ".xlsx": "Excel", ".xls": "Excel",
    ".pdf": "PDF", ".key": "Keynote", ".pages": "Pages", ".numbers": "Numbers",
    ".odt": "OpenDocument", ".rtf": "RTF",
}


def convert_text_file(src, out, convert):
    if out is None:
        base, ext = os.path.splitext(src)
        out = base + ".s2hk" + ext
    ext = os.path.splitext(src)[1].lower()
    if ext in BINARY_HINTS:
        sys.stderr.write(
            "ERROR: %s is a %s file, not plain text — this script cannot convert it.\n"
            "       Supported: inline --text, plain text (.txt/.md/.csv/.srt), and .pptx.\n"
            "       For %s, export/convert the content to text first, or use a dedicated tool.\n"
            % (src, BINARY_HINTS[ext], BINARY_HINTS[ext])
        )
        sys.exit(2)
    try:
        with open(src, "r", encoding="utf-8") as f:
            data = f.read()
    except UnicodeDecodeError:
        sys.stderr.write(
            "ERROR: %s is not UTF-8 plain text (looks like a binary or differently-encoded file).\n"
            "       Supported: inline --text, plain text (.txt/.md/.csv/.srt), and .pptx.\n"
            "       If it is GBK/Big5 text, re-save it as UTF-8 first.\n" % src
        )
        sys.exit(2)
    converted = convert(data)
    with open(out, "w", encoding="utf-8") as f:
        f.write(converted)
    changed = sum(1 for a, b in zip(data, converted) if a != b)
    print("text file converted")
    print("  input : %s" % src)
    print("  output: %s" % out)
    print("  chars changed (approx): %d" % changed)
    return out


def make_font_setter(etree, ea, latin):
    """返回把 <a:rPr> 的字体改写成指定值的函数；都为 None 时返回 None。

    只改 typeface 属性，保留 panose/charset 等其余属性。
    缺失的 latin/ea/cs 元素会被创建，并插到 <a:sym> 之前以维持合法子元素顺序。
    """
    if not ea and not latin:
        return None

    def set_font(rpr):
        def upsert(tagname, typeface):
            if typeface is None:
                return
            el = rpr.find(q(tagname))
            if el is None:
                el = etree.SubElement(rpr, q(tagname))
                el.set("typeface", typeface)
                sym = rpr.find(q("sym"))
                if sym is not None:
                    sym.addprevious(el)
            else:
                el.set("typeface", typeface)
        upsert("latin", latin)
        upsert("ea", ea)
        upsert("cs", latin)

    return set_font


def convert_pptx_file(src, out, convert, scopes=None, verbose=True,
                      ea=None, latin=None, font_scopes=None):
    import shutil
    import zipfile
    import tempfile
    from lxml import etree

    if out is None:
        base, ext = os.path.splitext(src)
        out = base + ".s2hk" + ext
    if scopes is None:
        scopes = list(SCOPE_ORDER)
    if font_scopes is None:
        font_scopes = ["slides"]

    set_font = make_font_setter(etree, ea, latin)

    def text_scope_of(name):
        if "/_rels/" in name:
            return None
        for s in scopes:
            if SCOPES[s](name):
                return s
        return None

    tmpdir = tempfile.mkdtemp()
    stats = dict.fromkeys(SCOPE_ORDER, 0)
    changed_text = 0
    fallback_paras = 0
    font_sets = 0
    try:
        with zipfile.ZipFile(src) as z:
            members = safe_members(z, tmpdir)
            infos = list(z.infolist())
            z.extractall(tmpdir, members=members)

        for name in [i.filename for i in infos]:
            scope = text_scope_of(name) if convert is not None else None
            do_font = set_font is not None and in_scope(name, font_scopes)
            if scope is None and not do_font:
                continue
            path = os.path.join(tmpdir, name)
            if not os.path.isfile(path):
                continue
            try:
                tree = etree.parse(path)
            except etree.XMLSyntaxError:
                continue
            root = tree.getroot()
            dirty = False

            if scope == "props":
                n = convert_props(root, convert,
                                  lambda el: etree.QName(el).localname)
                if n:
                    changed_text += n
                    stats[scope] += 1
                    dirty = True
            elif scope is not None:
                n, fb = convert_paragraph_aware(root, convert)
                fallback_paras += fb
                stats[scope] += 1
                if n:
                    changed_text += n
                    dirty = True

            if do_font:
                for tag in ("rPr", "defRPr", "endParaRPr"):
                    for rpr in root.iter(q(tag)):
                        set_font(rpr)
                        font_sets += 1
                        dirty = True

            if dirty:
                tree.write(path, xml_declaration=True, encoding="UTF-8",
                           standalone=True)

        repack(src, out, tmpdir, infos)
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)

    if verbose:
        print("pptx converted (layout/images/tables preserved)")
        print("  input : %s" % src)
        print("  output: %s" % out)
        if convert is not None:
            print("  parts processed   : %s" % (", ".join(
                "%s=%d" % (s, stats[s]) for s in SCOPE_ORDER if stats[s]) or "none"))
            print("  text nodes changed: %d" % changed_text)
        if set_font is not None:
            print("  font props set    : %d  (scope: %s)"
                  % (font_sets, ",".join(font_scopes)))
            if font_sets == 0:
                print("  note: 0 font props set — runs without any explicit "
                      "character formatting have no <a:rPr> to write into; "
                      "they inherit the theme font. This is not an error.")
        if fallback_paras:
            print("  note: %d paragraph(s) fell back to per-run conversion "
                  "(length changed after conversion)" % fallback_paras)
    return out


def main():
    ap = argparse.ArgumentParser(
        description="Simplified→Traditional (Hong Kong by default) converter using OpenCC."
    )
    ap.add_argument("input", nargs="?",
                    help="Input file path (.txt/.md/... or .pptx). Omit when using --text.")
    ap.add_argument("-o", "--output", default=None,
                    help="Output path. Defaults to <name>.s2hk<ext> next to the input.")
    ap.add_argument("--text", default=None,
                    help="Convert this inline string and print the result to stdout.")
    ap.add_argument("--mode", default="s2hk",
                    choices=["s2hk", "s2t", "s2tw", "s2twp",
                             "t2s", "hk2s", "tw2s", "tw2sp", "none"],
                    help="OpenCC conversion mode. Default: s2hk (Hong Kong Traditional). "
                         "Use 'none' for a font-only pass (requires --ea/--latin).")
    ap.add_argument("--via-simplified", action="store_true",
                    help="First normalize existing Traditional variants back to Simplified "
                         "(t2s), then convert to --mode. Use to fix wrong/mixed Traditional "
                         "(e.g. 戶->户, 纔->才) into consistent HK forms.")
    ap.add_argument("--scope", default="all",
                    help="PPTX only. Which parts to convert. 'all' (default) covers "
                         "slides,layouts,masters,notes,diagrams,charts,props. "
                         "'slides' = visible slides only (old behaviour). "
                         "Or a comma list, e.g. --scope slides,notes. "
                         "NOTE: 'diagrams' (SmartArt) matters — SmartArt text lives in "
                         "diagrams/data*.xml too, and editing the shape in PowerPoint "
                         "re-renders from there, reverting slide-only conversions.")
    ap.add_argument("--ea", default=None, metavar="FONT",
                    help="PPTX only. Chinese/East-Asian font name to apply, e.g. 楷体. "
                         "Omit to leave fonts untouched.")
    ap.add_argument("--latin", default=None, metavar="FONT",
                    help="PPTX only. Latin/English font name to apply (also used for "
                         "<a:cs>), e.g. Arial. Omit to leave fonts untouched.")
    ap.add_argument("--font-scope", default="slides",
                    help="Which parts to apply --ea/--latin to. Default 'slides' — "
                         "restyling masters/layouts changes the whole template's look, "
                         "which is rarely what was asked for. Use 'all' to include them.")
    ap.add_argument("--verify", action="store_true",
                    help="PPTX only. After converting, check structure, font profile "
                         "and leftover Simplified text (re-convert method), then report.")
    ap.add_argument("--dry-run", action="store_true",
                    help="Report what would change without writing the output file.")
    ap.add_argument("--json", action="store_true",
                    help="Emit a machine-readable JSON summary on stdout.")
    args = ap.parse_args()

    if args.mode == "none" and not (args.ea or args.latin):
        ap.error("--mode none 表示不转换文字，此时必须给出 --ea 或 --latin")

    if args.scope.strip().lower() == "all":
        scopes = list(SCOPE_ORDER)
    else:
        scopes = parse_scopes(args.scope, parser_error=ap.error)
    font_scopes = parse_scopes(args.font_scope, parser_error=ap.error)

    convert = None
    if args.mode != "none":
        try:
            convert = build_converter(args.mode, args.via_simplified)
        except ImportError:
            sys.stderr.write("ERROR: opencc not installed. Run: pip install opencc\n")
            sys.exit(1)

    if args.text is not None:
        convert_inline_text(args.text, convert)
        return

    if not args.input:
        ap.error("provide an input file path, or use --text \"...\"")

    if not os.path.isfile(args.input):
        sys.stderr.write("ERROR: file not found: %s\n" % args.input)
        sys.exit(1)

    ext = os.path.splitext(args.input)[1].lower()
    if ext == ".pptx":
        import tempfile
        import shutil
        real_out = args.output
        if real_out is None:
            base, e = os.path.splitext(args.input)
            real_out = base + ".s2hk" + e
        # dry-run 仍需真正跑一遍才能给出准确计数，只是写到临时文件后丢弃
        tmp_out = None
        if args.dry_run:
            tmp_out = os.path.join(tempfile.mkdtemp(), os.path.basename(real_out))
        try:
            written = convert_pptx_file(
                args.input, tmp_out or real_out, convert,
                scopes=scopes, verbose=not args.json,
                ea=args.ea, latin=args.latin, font_scopes=font_scopes)
        except ImportError:
            sys.stderr.write("ERROR: lxml required for .pptx. Run: pip install lxml\n")
            sys.exit(1)
        except ValueError as e:
            sys.stderr.write("ERROR: %s\n" % e)
            sys.exit(2)

        summary = {"input": args.input, "output": real_out, "mode": args.mode,
                   "scope": scopes, "dry_run": bool(args.dry_run)}
        exit_code = 0

        if args.verify and args.mode != "none":
            try:
                report, ok = verify(args.input, written, args.mode,
                                    expect_font_identical=not (args.ea or args.latin))
            except ImportError:
                sys.stderr.write("ERROR: --verify requires python-pptx. "
                                 "Run: pip install python-pptx\n")
                sys.exit(1)
            summary["verify"] = report
            summary["verify_ok"] = ok
            if not args.json:
                print("verify")
                print("  slides            : %d → %d" % (report["slides_src"], report["slides_out"]))
                print("  text runs         : %d → %d" % (report["runs_src"], report["runs_out"]))
                print("  font profile same : %s" % report["font_identical"])
                print("  media identical   : %s (%d files)"
                      % (report["media_identical"], report["media_count"]))
                print("  zip entry order   : %s" % report["entry_order_identical"])
                print("  residue (full doc): %d" % report["residue_count"])
                for part, text, diff in report["residue"][:10]:
                    print("    %s  %s  [%s]" % (part, text, diff))
                print("  result            : %s" % ("PASS" if ok else "FAIL"))
            if not ok:
                exit_code = 3

        if args.dry_run:
            if tmp_out:
                shutil.rmtree(os.path.dirname(tmp_out), ignore_errors=True)
            if not args.json:
                print("dry-run: no file written (would write %s)" % real_out)

        if args.json:
            import json
            json.dump(summary, sys.stdout, ensure_ascii=False, indent=2)
            sys.stdout.write("\n")
        sys.exit(exit_code)
    else:
        convert_text_file(args.input, args.output, convert)


if __name__ == "__main__":
    main()
