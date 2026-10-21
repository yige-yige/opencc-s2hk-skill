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

A = "http://schemas.openxmlformats.org/drawingml/2006/main"


def q(tag):
    return "{%s}%s" % (A, tag)


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


def convert_pptx_file(src, out, convert):
    import shutil
    import zipfile
    import tempfile
    from lxml import etree

    if out is None:
        base, ext = os.path.splitext(src)
        out = base + ".s2hk" + ext

    tmpdir = tempfile.mkdtemp()
    try:
        with zipfile.ZipFile(src) as z:
            names = z.namelist()
            z.extractall(tmpdir)

        slide_files = [
            n for n in names
            if n.startswith("ppt/slides/slide")
            and n.endswith(".xml")
            and "/_rels/" not in n
        ]

        changed_text = 0
        for name in slide_files:
            path = os.path.join(tmpdir, name)
            tree = etree.parse(path)
            root = tree.getroot()
            for t in root.iter(q("t")):
                if t.text:
                    new = convert(t.text)
                    if new != t.text:
                        changed_text += 1
                    t.text = new
            tree.write(path, xml_declaration=True, encoding="UTF-8", standalone=True)

        if os.path.exists(out):
            os.remove(out)
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zf:
            for folder, _, files in os.walk(tmpdir):
                for f in files:
                    full = os.path.join(folder, f)
                    arc = os.path.relpath(full, tmpdir)
                    zf.write(full, arc)
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)

    print("pptx converted (fonts/layout/images/tables preserved)")
    print("  input : %s" % src)
    print("  output: %s" % out)
    print("  slides processed  : %d" % len(slide_files))
    print("  text nodes changed: %d" % changed_text)
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
                             "t2s", "hk2s", "tw2s", "tw2sp"],
                    help="OpenCC conversion mode. Default: s2hk (Hong Kong Traditional).")
    ap.add_argument("--via-simplified", action="store_true",
                    help="First normalize existing Traditional variants back to Simplified "
                         "(t2s), then convert to --mode. Use to fix wrong/mixed Traditional "
                         "(e.g. 戶->户, 纔->才) into consistent HK forms.")
    args = ap.parse_args()

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
        try:
            convert_pptx_file(args.input, args.output, convert)
        except ImportError:
            sys.stderr.write("ERROR: lxml required for .pptx. Run: pip install lxml\n")
            sys.exit(1)
    else:
        convert_text_file(args.input, args.output, convert)


if __name__ == "__main__":
    main()
