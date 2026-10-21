import os
import sys


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


_bootstrap(("opencc", "pptx"))

USAGE = ("usage: verify.py <source.pptx> <output.pptx> <mode>\n"
         "  NOTE: 3 POSITIONAL args. There is NO --mode flag.\n"
         "  correct : verify.py in.pptx out.pptx s2hk\n"
         "  wrong   : verify.py in.pptx out.pptx --mode s2hk\n")

if len(sys.argv) != 4:
    sys.stderr.write(USAGE)
    sys.exit(2)
if sys.argv[3].startswith('-'):
    sys.stderr.write("ERROR: mode must be a bare value like 's2hk', not '%s'.\n\n%s"
                     % (sys.argv[3], USAGE))
    sys.exit(2)

try:
    from pptx import Presentation
    from opencc import OpenCC
except ImportError as e:
    sys.stderr.write(
        "ERROR: missing dependency (%s).\n"
        "       Install with:  %s -m pip install python-pptx opencc\n"
        "       If you hit 'externally-managed-environment', add --user.\n"
        % (getattr(e, 'name', e), sys.executable)
    )
    sys.exit(1)

A='http://schemas.openxmlformats.org/drawingml/2006/main'
RPR='{%s}rPr'%A; LATIN='{%s}latin'%A; EA='{%s}ea'%A
src,out,mode=sys.argv[1],sys.argv[2],sys.argv[3]
def collect(shapes,acc):
    for sh in shapes:
        if sh.shape_type==6: collect(sh.shapes,acc)
        if sh.has_text_frame:
            for p in sh.text_frame.paragraphs:
                for r in p.runs: acc.append(r)
        if sh.has_table:
            for row in sh.table.rows:
                for c in row.cells:
                    for p in c.text_frame.paragraphs:
                        for r in p.runs: acc.append(r)
    return acc
def profile(fn):
    prof={}
    for s in Presentation(fn).slides:
        for r in collect(s.shapes,[]):
            if not r.text.strip(): continue
            rpr=r._r.find(RPR)
            if rpr is None: continue
            la=rpr.find(LATIN); ea=rpr.find(EA)
            lav=la.get('typeface') if la is not None else None
            eav=ea.get('typeface') if ea is not None else None
            prof[(lav,eav)]=prof.get((lav,eav),0)+1
    return prof
print('font identical:', profile(src)==profile(out))
prs=Presentation(out)
print('visible slides:', len(prs.slides))
cc=OpenCC(mode); resid=[]
for i,s in enumerate(prs.slides,1):
    for r in collect(s.shapes,[]):
        t=r.text
        if t.strip() and cc.convert(t)!=t:
            conv=cc.convert(t)
            diff=''.join('%s->%s '%(x,y) for x,y in zip(t,conv) if x!=y)
            resid.append((i,diff.strip(),t.strip()[:40]))
print('residue warnings:', len(resid))
[print('  p%d [%s]  %s'%(i,d,txt)) for i,d,txt in resid]