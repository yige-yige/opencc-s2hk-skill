---
name: opencc-s2hk
description: Convert Chinese text from Simplified to Hong Kong Traditional (香港繁體, s2hk) — or any OpenCC variant (s2t / s2tw / s2twp / t2s / hk2s) — for inline strings and plain-text files (.txt / .md / .csv / .srt); also handles .pptx as a text-only pass that keeps layout and fonts byte-identical. Use when the user says 简体转繁体, 转香港繁体, 港式用字, 繁简转换, s2hk, OpenCC, or needs to regularize already-Traditional text that mixes 戶/纔/裡/啟 into consistent HK forms. If the task is a .pptx AND involves changing fonts (统一中英文字体 / 中文改楷体 / 英文改Arial), use the `pptx-s2t-font-convert` skill instead — this skill has no font capability.
description_zh: 简体转香港繁体（文本/Markdown，兼顾 PPTX）
description_en: Simplified to HK Traditional converter (text / md / pptx)
license: MIT
---

# opencc-s2hk

把中文由简体转为**香港繁体（s2hk 香港用字标准）**，也支持其它 OpenCC 变体。基于 OpenCC 字形转换，只改文字、不动样式。

## 何时使用

- 「把这段话转成香港繁体 / 港式用字」
- 「这个 PPT 转繁体，版式字体不要变」
- 「这份 Markdown / txt / srt 转繁体」
- 「这份文件已经是繁体，但 戶/纔/裡/啟 用字混乱，统一成香港写法」

## 与 `pptx-s2t-font-convert` 的分工（先看这里）

两个 skill 都用 opencc、都能处理 .pptx。**按输入类型分：**

| 用户的输入 | 用哪个 |
|---|---|
| 内联文本、`.txt` / `.md` / `.csv` / `.srt` | **本 skill**（对方读不了这些格式） |
| `.pptx` + 要改字体（统一中英文字体 / 改楷体） | **`pptx-s2t-font-convert`**（只有它能改字体） |
| `.pptx`，纯文字转换不碰字体 | 两个都行；对方多了 verify.py 和 LibreOffice 视觉校验 |

硬规则：涉及**字体**就转 `pptx-s2t-font-convert`，本 skill 没有字体能力。

## 前置依赖（首次使用必须确认）

本 skill 若通过 `install.sh` 安装，依赖**已由安装器装好**，正常无需再处理；脚本还会在系统 `python3` 缺依赖时，自动切到安装器记录在 `scripts/.python` 里的解释器（独立 venv）。只有在下述检查失败时才需要手动装：

```bash
pip install opencc      # 文本转换必需
pip install lxml        # 仅处理 .pptx 时需要
python3 -c "import opencc; print('opencc ok')"
```

若 `import opencc` 报 ModuleNotFoundError，先装依赖再继续，不要改用正则或手写字表替代。
遇到 `externally-managed-environment` 时改用 `python3 -m pip install --user --break-system-packages opencc lxml`。

## 用法

脚本位于本 skill 的 `scripts/s2hk_convert.py`（用它的**绝对路径**调用，不要写成 `@scripts/...`）。
若通过安装包安装，还有等价的全局命令可直接用，拿不准路径时优先用它：

```bash
~/.local/bin/s2hk --text "..."     # 等价于 python3 <skill>/scripts/s2hk_convert.py --text "..."
```

```bash
# 1) 内联文本 → 打印结果
python3 scripts/s2hk_convert.py --text "用户启用多租户数据隔离机制"
# → 用户啓用多租户數據隔離機制

# 2) 纯文本文件（默认输出 notes.s2hk.md）
python3 scripts/s2hk_convert.py notes.md
python3 scripts/s2hk_convert.py notes.md -o notes_hk.md

# 3) PPTX（默认输出 deck.s2hk.pptx，字体版式全保留）
python3 scripts/s2hk_convert.py deck.pptx
python3 scripts/s2hk_convert.py deck.pptx -o deck_hk.pptx

# 4) 已是繁体但用字混乱 → 规整为香港用字
python3 scripts/s2hk_convert.py already_trad.pptx --mode s2hk --via-simplified
```

## 转换模式 `--mode`

| mode | 方向 | 说明 |
|------|------|------|
| `s2hk`（默认） | 简 → 港繁 | 香港用字：裏面、啓用、覆盤、賬户、衞生（保留「户」不用「戶」） |
| `s2t` | 简 → 繁 | OpenCC 标准繁体 |
| `s2tw` | 简 → 台繁 | 台湾正体 |
| `s2twp` | 简 → 台繁 | 台湾正体 + 台湾用词替换 |
| `t2s` / `hk2s` / `tw2s` / `tw2sp` | 繁 → 简 | 反向转换 |

## `--via-simplified`

输入**已经是繁体**但字形混杂/不符合目标地区时使用：先 `t2s` 归一回简体，再转到目标 `--mode`，干净地统一为一致港式写法（`戶→户`、`纔→才`、`裡→裏`、`啟→啓`）。

## 香港 vs 台湾用字

- 香港：**裏**面、**啓**用、賬**户**、**衞**生
- 台湾：**裡**面、**啟**用、帳**戶**、**衛**生

面向香港政府 / 公众场景时用 s2hk。

## 验证转换是否彻底（重要）

**不要**用「固定简体字表」查残留 —— 香港繁体会合理保留一些看起来像简体的字（`户`、`峰`、`兌`），固定字表必然误报。

正确做法是**二次转换法**：对输出再跑一遍同一 mode，若 `convert(文本) == 文本` 对所有文本成立，即已彻底转换。

```python
from opencc import OpenCC
cc = OpenCC('s2hk')
residue = [t for t in all_output_texts if cc.convert(t) != t]
# len(residue) == 0 → 彻底
```

## 支持的输入形式（务必先确认）

| 形式 | 例子 | 支持 |
|---|---|---|
| 内联文本（直接给一句话，不涉及文件） | `--text "用户启用..."` | ✅ |
| UTF-8 纯文本文件 | `.txt` `.md` `.csv` `.srt` | ✅ |
| PPT | `.pptx` | ✅（纯文字，不改字体） |
| Office / 排版类二进制文件 | `.docx` `.xlsx` `.pdf` `.key` `.pages` | ❌ 脚本会报错退出（exit 2） |
| 非 UTF-8 文本（GBK / Big5） | — | ❌ 需先另存为 UTF-8 |

「纯文本文件」= 用记事本 / VS Code 打开就能看到原文的文件。`.docx` 虽然看起来是文档，实际是 zip 二进制容器，**不属于纯文本**。脚本已内置拦截并给出明确提示，不会再抛 `UnicodeDecodeError`。

## 已知边界

- PPTX 中**图片 / 截图 / 视频封面里的文字**属图像内容，无法编辑，保持原状。
- OpenCC 只做**字形**转换，不改用词习惯（「程序」不会自动变成港式「程式」）。用词本地化需另行处理。
- 脚本只处理 `ppt/slides/slideN.xml`；母版（slideMaster）、版式（slideLayout）、备注页（notesSlide）、图表内嵌文字**不在处理范围**。若用户的文字写在母版或备注里，需另行说明。
- `.docx` / `.pdf` / `.xlsx` 等二进制格式**不支持**，脚本会以 exit 2 报错。若用户拿来的是 Word，需先导出为 `.md` / `.txt`，或改用专门的 docx 工具链。
