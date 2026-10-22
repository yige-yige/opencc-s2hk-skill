#!/usr/bin/env python3
"""
Batch convert PPTX text language (S2T/T2S) and/or unify CN/EN fonts,
preserving layout, colors, images, tables, highlights and OLE objects.

Usage:
  python3 convert.py input.pptx output.pptx --ea "楷体" --latin "Arial" --mode s2t

--ea     : East-Asian (Chinese) font typeface name. Omit to leave CN font unchanged.
--latin  : Latin (English) font typeface name (also applied to <a:cs>). Omit to leave unchanged.
--mode   : s2t (Simplified->Traditional), t2s (Traditional->Simplified), none (no text conversion).

Only ppt/slides/slideN.xml is processed (covers grouped shapes & table cells).
OLE embedded object text is NOT touched (by design — kept as-is).
"""
import sys, os, shutil, zipfile, tempfile, argparse


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


_bootstrap(("opencc", "lxml", "pptx"))


def _missing(pkg, why):
    sys.stderr.write(
        "ERROR: %s is not installed (%s).\n"
        "       Install with:  %s -m pip install %s\n"
        "       If you hit 'externally-managed-environment', try:\n"
        "         %s -m pip install --user %s\n"
        % (pkg, why, sys.executable, pkg, sys.executable, pkg)
    )
    sys.exit(1)


try:
    from lxml import etree
except ImportError:
    _missing("lxml", "required for reading/writing pptx XML")

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
def q(tag): return f"{{{A}}}{tag}"

# PPTX 处理核心与本脚本同目录（skill 自包含，不跨目录 import）
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pptx_core import (  # noqa: E402
    SCOPES, SCOPE_ORDER, parse_scopes, safe_members,
    convert_paragraph_aware, repack,
)

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("src")
    ap.add_argument("out")
    ap.add_argument("--ea", default=None, help="Chinese/East-Asian font name")
    ap.add_argument("--latin", default=None, help="English/Latin font name (also used for <a:cs>)")
    ap.add_argument("--mode", default="s2t",
                    choices=["s2t", "t2s", "s2hk", "hk2s", "s2tw", "s2twp", "tw2s", "tw2sp", "none"],
                    help="Conversion mode. s2t=Simp->Trad(generic), s2hk=Simp->HongKong Trad, "
                         "s2tw=Simp->Taiwan Trad, s2twp=Simp->Taiwan Trad w/ phrases, "
                         "t2s/hk2s/tw2s = reverse. none=no text conversion.")
    ap.add_argument("--via-simplified", action="store_true",
                    help="Normalize any existing Traditional variants back to Simplified FIRST, "
                         "then convert to --mode target. Use when the source is already Traditional "
                         "but in a wrong/mixed variant (e.g. fixing 戶/纔 to HK 户/才 via s2hk). "
                         "Only meaningful for s2* target modes.")
    ap.add_argument("--scope", default="all",
                    help="Which parts to CONVERT TEXT in. 'all' (default) = slides,layouts,"
                         "masters,notes,diagrams,charts,props. 'slides' = visible slides only "
                         "(old behaviour). SmartArt text also lives in diagrams/data*.xml — "
                         "editing the shape in PowerPoint re-renders from there, so "
                         "slide-only conversion silently reverts.")
    ap.add_argument("--font-scope", default="slides",
                    help="Which parts to APPLY FONTS to. Default 'slides' — rewriting fonts on "
                         "masters/layouts changes the whole template's look, which is rarely "
                         "what the user asked for. Use 'all' to include them.")
    args = ap.parse_args()

    text_scopes = parse_scopes(args.scope, parser_error=ap.error)
    font_scopes = parse_scopes(args.font_scope, parser_error=ap.error)

    cc = None
    pre = None
    if args.mode != "none":
        try:
            from opencc import OpenCC
        except ImportError:
            _missing("opencc", "required for Chinese text conversion")
        cc = OpenCC(args.mode)
        if args.via_simplified:
            pre = OpenCC('t2s')  # collapse any Traditional variant to Simplified first

    def convert_text(s):
        if pre is not None:
            s = pre.convert(s)
        return cc.convert(s)

    def set_font(rpr):
        def upsert(tagname, typeface):
            if typeface is None:
                return
            el = rpr.find(q(tagname))
            if el is None:
                el = etree.SubElement(rpr, q(tagname))
                el.set("typeface", typeface)
                sym = rpr.find(q('sym'))
                if sym is not None:
                    sym.addprevious(el)
            else:
                el.set("typeface", typeface)
        upsert('latin', args.latin)
        upsert('ea', args.ea)
        upsert('cs', args.latin)

    tmpdir = tempfile.mkdtemp()
    with zipfile.ZipFile(args.src) as z:
        infos = list(z.infolist())
        names = [i.filename for i in infos]
        z.extractall(tmpdir, members=safe_members(z, tmpdir))

    def _in(name, scopes):
        if "/_rels/" in name:
            return False
        return any(SCOPES[s](name) for s in scopes)

    text_files = {n for n in names if _in(n, text_scopes)}
    # 字体只改选定范围（默认仅正文页）：改母版/版式会波及整份模板的观感
    font_files = {n for n in names if _in(n, font_scopes)}
    targets = [n for n in names if n in text_files | font_files]

    changed_text = 0
    fallback_paras = 0
    font_sets = 0
    for name in targets:
        path = os.path.join(tmpdir, name)
        if not os.path.isfile(path):
            continue
        try:
            tree = etree.parse(path)
        except etree.XMLSyntaxError:
            continue
        root = tree.getroot()
        if cc is not None and name in text_files:
            n, fb = convert_paragraph_aware(root, convert_text)
            changed_text += n
            fallback_paras += fb
        if (args.ea or args.latin) and name in font_files:
            for tag in ('rPr', 'defRPr', 'endParaRPr'):
                for rpr in root.iter(q(tag)):
                    set_font(rpr)
                    font_sets += 1
        tree.write(path, xml_declaration=True, encoding="UTF-8", standalone=True)

    # 保真重打包：沿用原压缩方式/顺序/时间戳，不把 STORED 条目重新压缩
    repack(args.src, args.out, tmpdir, infos)
    shutil.rmtree(tmpdir, ignore_errors=True)

    print(f"xml parts processed: {len(targets)}  "
          f"(text scope: {','.join(text_scopes)} | font scope: "
          f"{','.join(font_scopes) if (args.ea or args.latin) else 'n/a'})")
    print(f"text nodes changed: {changed_text}")
    if fallback_paras:
        print(f"note: {fallback_paras} paragraph(s) fell back to per-run conversion")
    print(f"font props set    : {font_sets}")
    print(f"output            : {args.out}")

if __name__ == "__main__":
    try:
        main()
    except ValueError as e:
        # 不可信输入（Zip Slip 等）：给出可读错误，不抛裸 traceback
        sys.stderr.write("ERROR: %s\n" % e)
        sys.exit(2)
