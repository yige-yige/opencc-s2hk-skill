---
name: pptx-s2t-font-convert
description: "**PPTX FILES ONLY.** Batch-convert all text in a .pptx from
  Simplified to Traditional Chinese (or vice versa) AND/OR unify Chinese vs
  English fonts, while keeping layout, colors, images, icons, highlights, tables
  and OLE objects fully intact. Use when the user wants to \"把PPT文字转繁体/简体\",
  \"统一PPT中英文字体\", \"中文改楷体/宋体 英文改Arial\", or any request to change PPTX text
  language conversion and/or per-language font without altering anything else.
  Adds font control + LibreOffice visual QA on top of plain conversion. NOT for
  plain text, Markdown, .txt/.csv/.srt or inline strings — for those use the
  `opencc-s2hk` skill instead. Operates by unpacking pptx, editing only
  ppt/slides/slideN.xml at the XML level (opencc on <a:t>, font typeface on rPr
  latin/ea/cs), then repacking — far safer than touching the design."
description_zh: PPT文字繁简转换与中英文字体统一
description_en: PPTX S2T/T2S text & font convert
license: MIT
---

# pptx-s2t-font-convert

A reusable workflow to convert all text in a PPTX between Simplified/Traditional Chinese and/or unify Chinese (East-Asian) and English (Latin) fonts, while preserving everything else (layout, colors, images, icons, highlights, tables).

## When to use
- "把这个 PPT 所有文字转成繁体/简体中文"
- "中文字体改成楷体，英文改成 Arial，其他不变"
- "统一 PPT 的中英文字体" / "PPT 字体规范化"
- Any combination: language conversion + per-language font, with strict "其他全保持原状" requirement.

## Relationship to `opencc-s2hk` (read this first — avoid picking the wrong skill)
Both skills convert Chinese via opencc and both handle .pptx. **Split by INPUT TYPE:**

| User's input | Use |
|---|---|
| `.pptx` | **this skill** (font control + visual QA + verify.py) |
| inline text, `.txt` / `.md` / `.csv` / `.srt` | **`opencc-s2hk`** (this skill cannot read them) |
| `.pptx` + "顺便统一字体" | **this skill** (only this one does fonts) |
| `.pptx`, text-only, no font change | either works; this one has stronger QA |

Hard rule: if the input is NOT a .pptx, stop and use `opencc-s2hk`. If the request involves fonts, stay here — `opencc-s2hk` has no font capability.

## Conversion mode / regional variants (IMPORTANT)
opencc supports regional variants — pick the mode that matches the user's target locale:
- `s2t` : Simplified → Traditional (generic/OpenCC standard). Default.
- `s2hk`: Simplified → **Hong Kong Traditional** (香港繁體). Use when user says "香港繁体 / 現代香港用字 / 港式". e.g. 裏面、啓動、覆盤、記憶體; keeps 户 (HK uses 户 not 戶).
- `s2tw`: Simplified → Taiwan Traditional (台灣正體). `s2twp` adds Taiwan-specific vocabulary substitution.
- `t2s` / `hk2s` / `tw2s` : reverse (→ Simplified). `tw2sp` includes phrase conversion.
- `none`: no text conversion (font-only pass).
When the user distinguishes "香港繁体 vs 传统正体/晦涩正体", they usually mean s2hk (NOT s2t/s2tw). 香港 → s2hk, 台湾 → s2tw/s2twp.

### Fixing an ALREADY-Traditional file into the right variant (`--via-simplified`)
When the source is already Traditional but in a WRONG/MIXED variant (e.g. filename says "繁体有错别字", or it has 戶/纔 that should be HK 户/才), a single s2hk pass over Traditional text is imperfect. Use `--via-simplified`: it first collapses ALL Traditional variants back to Simplified (t2s), then applies the target mode (e.g. s2hk). This cleanly regularizes 戶→户, 纔→才, 裡→裏, 啟→啓 into consistent HK forms.
```bash
convert.py in.pptx out.pptx --mode s2hk --via-simplified
```
Residue verification for this path: `s2hk(t2s(text)) == text` for every run.

## Key design decision
Do NOT rebuild slides or use python-pptx run-by-run editing (misses grouped shapes, table cells, empty runs). Instead edit the raw slide XML:
- Convert every `<a:t>` text node with opencc (`s2t`/`s2hk`/`s2tw` for 简→繁, `t2s` for 繁→简).
- For every `<a:rPr>`, `<a:defRPr>`, `<a:endParaRPr>`: set the `typeface` of `<a:latin>` and `<a:cs>` to the English font, and `<a:ea>` to the Chinese font. Only change `typeface`, keep panose/charset/etc. Create the element if missing (insert before `<a:sym>` if present).
- Only process `ppt/slides/slideN.xml` (the visible content). This covers grouped/nested shapes and table cells automatically since they are all in the same XML.

## Prerequisites
Required Python packages: `python-pptx`, `opencc`, `lxml`.

If this skill was installed via the bundled `install.sh`, dependencies are **already in place** — and both scripts self-heal: when the running `python3` lacks the deps, they re-exec with the interpreter recorded in `scripts/.python` (a dedicated venv created by the installer). So just run them; only fall back to manual installation if that fails.

**Do NOT hardcode a python path.** Resolve an interpreter first, then reuse it as `$PY`:
```bash
PY="$(command -v python3)"   # on this workspace, a managed env may be preferable
"$PY" -c "import pptx,opencc,lxml; print('deps ok')"
# if that fails:
"$PY" -m pip install python-pptx opencc lxml
# on 'externally-managed-environment' (newer macOS/Linux):
"$PY" -m pip install --user python-pptx opencc lxml
```
Both bundled scripts now fail with an actionable message (naming the exact pip command for the running interpreter) instead of a bare `ModuleNotFoundError`.

**Optional** — only for the visual-QA step: LibreOffice `soffice` + `pdftoppm`. If they are absent, conversion and verify.py still work fully; skip step 4 and say so rather than reporting a failure.

## Steps
1. Confirm with the user the exact Chinese font NAME to write (e.g. 楷体 / KaiTi / 標楷體 DFKai-SB / 宋体). The literal name matters — use what the user says; "楷体" is the safest cross-platform generic name for 楷体.
2. Run the conversion script. Paths below are relative to this skill's own directory — use the real path on disk (`$SKILL_DIR/scripts/convert.py`), **not** a literal `@scripts/...`, which is not a valid path and fails with `can't open file '.../@scripts/convert.py'`:
   ```bash
   "$PY" "$SKILL_DIR/scripts/convert.py" "input.pptx" "output.pptx" --ea "楷体" --latin "Arial" --mode s2t
   ```
   (Or inline the logic — see scripts/convert.py for the canonical implementation.)
3. Verify with the bundled QA script (recommended). **verify.py takes 3 POSITIONAL args — there is NO `--mode` flag**; passing `--mode s2hk` used to make OpenCC raise `RuntimeError: --mode not found or not accessible`, and now exits 2 with a usage hint:
   ```bash
   "$PY" "$SKILL_DIR/scripts/verify.py" "input.pptx" "output.pptx" s2hk   # correct
   # NOT: ... "output.pptx" --mode s2hk   ← rejected
   ```
   → prints `font identical` (**expected True ONLY on a font-unchanged pass**; if you intentionally passed `--ea`/`--latin`, False is correct and expected — do not treat it as a failure), `visible slides` (= len(prs.slides)), and `residue warnings` (0 = clean; 1-2 may be OpenCC false positives, inspect). Or verify manually (see Verification below).
4. **Optional, requires soffice + pdftoppm** — render to images and do a visual QA pass on 2-3 representative slides (a title page, a text-heavy Chinese page). Skip if the tools are missing.
5. Deliver the output file to the user (e.g. via `present_files`).

## Pitfalls
- **OLE embedded objects** (shape_type EMBEDDED_OLE_OBJECT) hold text inside an embedded document — python-pptx/XML edits CANNOT reach it. Under "保持原状" this is correct (leave as-is), but tell the user these stay unchanged.
- **markitdown does not extract OLE text** — so a clean grep for Simplified chars does not prove OLE text was converted (it wasn't, by design).
- When creating a missing `<a:latin/ea/cs>` element, insert it before `<a:sym>` to keep valid child order; appending at the end also generally works but ordering is safer.
- LibreOffice may lack the exact font (e.g. 楷体) and render a fallback — that's a rendering-environment issue, NOT a file problem. The typeface name in the file is what matters; it renders correctly on machines that have the font.
- Repack must preserve the zip structure; walk the temp dir and re-zip with ZIP_DEFLATED. Keep `[Content_Types].xml` and rels untouched.
- Always write a NEW output file, never overwrite the source.
- **`font props set: 0` is NOT necessarily a failure.** Fonts are written by upserting into EXISTING `<a:rPr>` / `<a:defRPr>` / `<a:endParaRPr>` elements. A run that never had any explicit character formatting has NO `rPr` at all, so there is nothing to upsert into and it silently inherits the master/theme font. Verified: a deck with 0 `rPr` reports `font props set: 0` and keeps its original fonts; after giving one run an explicit size (creating an `rPr`), the same command reports `font props set: 1` and correctly writes `<a:ea typeface="楷体"/>` + `<a:latin typeface="Arial"/>`. If the user demands EVERY run change font, the script must be extended to create `rPr` on bare runs — flag this instead of silently reporting success.
- **Scope is configurable now (`--scope` / `--font-scope`).** Text conversion defaults to
  `all` = slides + slideLayouts + slideMasters/notesMasters + notesSlides + diagrams
  (SmartArt) + charts + docProps. **Font** rewriting defaults to `slides` only, because
  restyling masters/layouts changes the whole template's look — pass `--font-scope all`
  only if the user explicitly wants that. SmartArt matters: its text also lives in
  `ppt/diagrams/data*.xml`, and editing the shape in PowerPoint re-renders from there,
  so a slides-only pass silently reverts.
- **"slides processed" counts slideN.xml files, NOT the deck's visible page count.** A pptx can carry orphaned/unreferenced slide XML (not in presentation.xml sldIdLst), so the number may exceed the real page count (e.g. reports 5 but the deck shows 2). Processing all of them is intentional/safe. To get the TRUE visible page count, open with python-pptx and use `len(prs.slides)`. Don't be alarmed by the mismatch; just verify visible pages in QA.

## Verification
- Reopen with python-pptx (no exception = structurally valid).
- Sample Chinese runs and confirm `latin/ea/cs` typeface values are as intended.
- **Residual-conversion check — use the "re-convert" method, NOT a fixed simplified-char dictionary.**
  Re-run the SAME opencc mode over the output text; if `cc.convert(t) == t` for every run, conversion is complete (diff count = 0 means no leftovers).
  ```python
  from opencc import OpenCC
  cc = OpenCC('s2hk')  # same mode as used
  diff = [t for t in all_output_texts if cc.convert(t) != t]
  # len(diff) == 0  → fully converted
  ```
  WHY: a hardcoded 简体字 set gives false positives for s2hk/s2tw because HK/TW legitimately KEEP some chars that look "simplified" — e.g. HK keeps 户 (not 戶), 峰, 兌. Don't flag those as residue. The re-convert method is variant-aware and reliable.
  CAVEAT (rare false positives): the re-convert method can itself flag 1-2 items due to OpenCC's phrase-level segmentation, where an ALREADY-correct word gets "re-converted" to a rarer variant. Observed false-positive list: 適合→適閤 (in 更適合中大型), 團伙→團夥. These are NOT real residue. Decision rule: run `OpenCC(mode).convert(<the SIMPLIFIED source word>)` — whatever THAT outputs is the standard (s2hk('团伙')='團伙', s2hk('适合')='適合'); the re-convert only "wants" to change it because feeding already-Traditional text back triggers a different phrase rule. Keep the first-pass form, do NOT "fix". Always confirm via visual QA.
- Font-preservation check: compare (latin, ea) typeface distribution of source vs output — must be identical when doing a font-only-unchanged pass.
- Visual QA: `soffice --headless --convert-to pdf` then `pdftoppm -jpeg -r 110`, read 2-3 slides to confirm target-variant 繁/简 + correct fonts + layout preserved. For s2hk, spot-check HK-specific forms (裏/啓/覆/記憶體) rather than TW forms (裡/啟).
