#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""pptx_core — 两个 skill 共用的 PPTX 处理核心。

集中实现三件容易写错、且必须保持一致的事：
  1. 安全解压   —— 拦截 Zip Slip（恶意 .pptx 用 ../ 把文件写到目标目录外）
  2. 处理范围   —— 只改 ppt/slides/ 会漏掉版式/母版/备注/SmartArt 里的文字
  3. 段落级转换 —— PPT 常把一个词拆进多个 <a:t>，逐节点转换会破坏词组规则

这个文件被两个 skill 各存一份（skill 必须自包含，不能跨目录 import）。
修改时请同步另一份，CI 会校验两份内容一致。
"""
import os
import zipfile

A = "http://schemas.openxmlformats.org/drawingml/2006/main"


def q(tag):
    return "{%s}%s" % (A, tag)


# --------------------------------------------------------------------------
# 处理范围
#
# SmartArt 尤其值得注意：文字同时存在于 ppt/diagrams/data*.xml（数据源）和
# slide 里的缓存绘图中。只改缓存的话，用户在 PowerPoint 里一编辑该图形，
# 就会从 data.xml 重新渲染，转换结果被打回原形。
SCOPES = {
    "slides": lambda n: n.startswith("ppt/slides/slide") and n.endswith(".xml"),
    "layouts": lambda n: n.startswith("ppt/slideLayouts/slideLayout") and n.endswith(".xml"),
    "masters": lambda n: (
        (n.startswith("ppt/slideMasters/slideMaster")
         or n.startswith("ppt/notesMasters/notesMaster")
         or n.startswith("ppt/handoutMasters/handoutMaster"))
        and n.endswith(".xml")),
    "notes": lambda n: n.startswith("ppt/notesSlides/notesSlide") and n.endswith(".xml"),
    "diagrams": lambda n: n.startswith("ppt/diagrams/") and n.endswith(".xml"),
    "charts": lambda n: n.startswith("ppt/charts/") and n.endswith(".xml"),
    "props": lambda n: n in ("docProps/core.xml", "docProps/app.xml"),
}
SCOPE_ORDER = ["slides", "layouts", "masters", "notes", "diagrams", "charts", "props"]

# 文档属性里需要转换的字段（按 localname 匹配，避免命名空间差异）
PROP_LOCALNAMES = {"title", "subject", "keywords", "description", "category"}


def parse_scopes(value, parser_error=None):
    """把 --scope 的取值解析成 scope 列表。"""
    value = (value or "all").strip().lower()
    if value == "all":
        return list(SCOPE_ORDER)
    scopes = [s.strip() for s in value.split(",") if s.strip()]
    bad = [s for s in scopes if s not in SCOPES]
    if bad:
        msg = ("--scope 不支持 %s；可选：all 或 %s 的逗号组合"
               % (", ".join(bad), ",".join(SCOPE_ORDER)))
        if parser_error:
            parser_error(msg)
        raise ValueError(msg)
    return scopes


def in_scope(name, scopes):
    if "/_rels/" in name:
        return False
    return any(SCOPES[s](name) for s in scopes)


def safe_members(zf, dest):
    """返回可安全解压的成员，拦截 Zip Slip。

    .pptx 多来自用户或网络，属不可信输入；标准库的 extractall 并不保证
    阻止恶意条目名把文件写到目标目录之外，必须显式校验。
    """
    dest_real = os.path.realpath(dest)
    safe = []
    for info in zf.infolist():
        name = info.filename
        if name.endswith("/"):
            continue
        if os.path.isabs(name) or name.startswith(("/", "\\")) or ":" in name[:2]:
            raise ValueError("压缩包含绝对路径条目，已拒绝：%s" % name)
        target = os.path.realpath(os.path.join(dest, name))
        if target != dest_real and not target.startswith(dest_real + os.sep):
            raise ValueError("压缩包含路径穿越条目，已拒绝：%s" % name)
        safe.append(info)
    return safe


def convert_paragraph_aware(root, convert):
    """按段落聚合后再转换，避免词组被 run 边界切断。

    PPT 常把一个段落拆成多个 <a:t>（加粗半个词、插入空 run 等），
    逐节点转换会让 OpenCC 的词组规则失效：
        '头发' 整段转 → 頭髮 ；拆成 '头'+'发' 逐个转 → 頭發（错）
    这里先拼接整段、转换一次，再按原长度回填。
    若转换后长度变化（s2twp 等用词替换会改变字数），该段回退为逐节点
    转换以保证不串位，并计入 fallback 供调用方提示。

    返回 (changed_nodes, fallback_paragraphs)。
    """
    changed = 0
    fallback = 0
    handled = set()

    for p in root.iter(q("p")):
        nodes = [t for t in p.iter(q("t")) if t.text]
        if not nodes:
            continue
        for t in nodes:
            handled.add(id(t))

        if len(nodes) == 1:
            new = convert(nodes[0].text)
            if new != nodes[0].text:
                nodes[0].text = new
                changed += 1
            continue

        joined = "".join(t.text for t in nodes)
        converted = convert(joined)
        if len(converted) == len(joined):
            pos = 0
            for t in nodes:
                seg = converted[pos:pos + len(t.text)]
                pos += len(t.text)
                if seg != t.text:
                    t.text = seg
                    changed += 1
        else:
            fallback += 1
            for t in nodes:
                new = convert(t.text)
                if new != t.text:
                    t.text = new
                    changed += 1

    # 段落之外的游离文本节点（图表标签等）
    for t in root.iter(q("t")):
        if id(t) in handled or not t.text:
            continue
        new = convert(t.text)
        if new != t.text:
            t.text = new
            changed += 1

    return changed, fallback


def convert_props(root, convert, qname_localname):
    """转换 docProps 中的标题/主题等字段。"""
    n = 0
    for el in root.iter():
        if qname_localname(el) in PROP_LOCALNAMES and el.text:
            new = convert(el.text)
            if new != el.text:
                el.text = new
                n += 1
    return n


def repack(src, out, tmpdir, infos):
    """保真重打包：沿用原压缩方式、顺序与时间戳。

    直接 os.walk + ZIP_DEFLATED 会把原本 STORED（不压缩）的条目重新压缩，
    对已压缩的 PNG/JPEG 既不省空间又损失保真度。
    """
    if os.path.exists(out):
        os.remove(out)
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(out, "w") as zf:
        for info in infos:
            if info.filename.endswith("/"):
                zf.writestr(info, b"")
                continue
            path = os.path.join(tmpdir, info.filename)
            if os.path.isfile(path):
                with open(path, "rb") as f:
                    data = f.read()
            else:
                data = zin.read(info.filename)
            ni = zipfile.ZipInfo(info.filename, date_time=info.date_time)
            ni.compress_type = info.compress_type
            ni.external_attr = info.external_attr
            ni.internal_attr = info.internal_attr
            ni.create_system = info.create_system
            zf.writestr(ni, data)
    return out


# --------------------------------------------------------------------------
# 质检
#
# 残留检测必须用「二次转换法」，不能用固定简体字表：香港繁体会合理保留一些
# 看起来像简体的字（户、峰、兌），固定字表必然误报。
def _iter_runs(shapes, acc):
    for sh in shapes:
        if sh.shape_type == 6:          # GROUP
            _iter_runs(sh.shapes, acc)
        if getattr(sh, "has_text_frame", False):
            for p in sh.text_frame.paragraphs:
                for r in p.runs:
                    acc.append(r)
        if getattr(sh, "has_table", False):
            for row in sh.table.rows:
                for c in row.cells:
                    for p in c.text_frame.paragraphs:
                        for r in p.runs:
                            acc.append(r)
    return acc


def font_profile(path):
    """统计 (latin, ea) 字体组合的分布，用于比对字体是否被意外改动。"""
    from pptx import Presentation
    rpr = q("rPr")
    latin, ea = q("latin"), q("ea")
    prof = {}
    for s in Presentation(path).slides:
        for r in _iter_runs(s.shapes, []):
            if not r.text.strip():
                continue
            el = r._r.find(rpr)
            if el is None:
                continue
            la, e = el.find(latin), el.find(ea)
            key = (la.get("typeface") if la is not None else None,
                   e.get("typeface") if e is not None else None)
            prof[key] = prof.get(key, 0) + 1
    return prof


def full_text_residue(path, mode):
    """全文（含版式/母版/备注/SmartArt）残留检查，返回 [(部件, 文本, 差异)]。"""
    from lxml import etree
    from opencc import OpenCC
    cc = OpenCC(mode)
    out = []
    with zipfile.ZipFile(path) as z:
        for name in z.namelist():
            if not name.endswith(".xml") or "/_rels/" in name:
                continue
            try:
                root = etree.fromstring(z.read(name))
            except Exception:
                continue
            for t in root.iter(q("t")):
                if not t.text or not t.text.strip():
                    continue
                conv = cc.convert(t.text)
                if conv != t.text:
                    diff = " ".join("%s→%s" % (a, b)
                                    for a, b in zip(t.text, conv) if a != b)
                    out.append((name, t.text.strip()[:40], diff))
    return out


def verify(src, out, mode, expect_font_identical=True):
    """对转换结果做结构/字体/残留三项质检，返回 (report_dict, ok)。"""
    from pptx import Presentation
    report = {}
    ok = True

    ps, po = Presentation(src), Presentation(out)
    report["slides_src"] = len(ps.slides)
    report["slides_out"] = len(po.slides)
    if report["slides_src"] != report["slides_out"]:
        ok = False

    runs_src = sum(len(_iter_runs(s.shapes, [])) for s in ps.slides)
    runs_out = sum(len(_iter_runs(s.shapes, [])) for s in po.slides)
    report["runs_src"], report["runs_out"] = runs_src, runs_out
    if runs_src != runs_out:
        ok = False

    identical = font_profile(src) == font_profile(out)
    report["font_identical"] = identical
    # 只有在「本次没打算改字体」时，字体变化才算问题
    if expect_font_identical and not identical:
        ok = False

    with zipfile.ZipFile(src) as a, zipfile.ZipFile(out) as b:
        media = [n for n in a.namelist() if n.startswith("ppt/media/")]
        same = (set(media) <= set(b.namelist())
                and all(a.read(n) == b.read(n) for n in media))
        report["media_count"] = len(media)
        report["media_identical"] = same
        report["entry_order_identical"] = a.namelist() == b.namelist()
    if not same:
        ok = False

    resid = full_text_residue(out, mode)
    report["residue"] = resid
    report["residue_count"] = len(resid)
    if resid:
        ok = False

    return report, ok
