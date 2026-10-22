---
name: opencc-s2hk
description: 简繁转换（默认香港繁体 s2hk）。处理内联文本、.txt/.md/.csv/.srt，以及 .pptx —— PPT 覆盖正文页/版式/母版/备注/SmartArt，可选统一中英文字体，版式图片表格全保留。触发词：简体转繁体、转香港繁体、港式用字、繁简转换、PPT转繁体、统一中英文字体、中文改楷体、s2hk、OpenCC；也用于把已是繁体但 戶/纔/裡/啟 混用的文档规整为一致港式用字。
description_zh: 简繁转换与 PPT 字体统一（默认香港繁体）
description_en: Simplified/Traditional converter with PPTX font control
license: MIT
---

# opencc-s2hk

把中文在简体与繁体之间转换，**默认输出香港繁体（s2hk 香港用字标准）**。基于 OpenCC 做字形转换；处理 PPTX 时在 XML 层只改文字与（可选的）字体名，版式、配色、图片、表格原样保留。

## 何时使用

- 「把这段话 / 这份 Markdown 转成香港繁体」
- 「这个 PPT 转繁体，版式不要变」
- 「PPT 转繁体，顺便中文改楷体、英文改 Arial」
- 「这份文件已是繁体，但 戶/纔/裡/啟 用字混乱，统一成香港写法」

## 前置依赖

经 `install.sh` 安装时依赖已就绪；脚本还会在当前 `python3` 缺依赖时自动切到安装器记录的解释器。手动安装：

```bash
python3 -m pip install opencc lxml        # lxml 仅 .pptx 需要
python3 -m pip install python-pptx        # 仅 --verify 需要
```

报 `externally-managed-environment` 时加 `--user --break-system-packages`。
若 `import opencc` 失败，先装依赖，**不要**改用正则或手写字表替代。

## 用法

脚本在本 skill 的 `scripts/s2hk_convert.py`（用绝对路径调用）。装过安装包的话有等价全局命令 `s2hk`，拿不准路径时优先用它。

```bash
# 内联文本 → 打印结果
s2hk --text "用户启用多租户数据隔离机制"

# 纯文本文件（默认输出 notes.s2hk.md）
s2hk notes.md
s2hk notes.md -o notes_hk.md

# PPTX（默认输出 deck.s2hk.pptx）
s2hk deck.pptx
s2hk deck.pptx -o deck_hk.pptx --mode s2hk

# PPTX + 统一字体
s2hk deck.pptx -o out.pptx --ea "楷体" --latin "Arial"

# 只改字体、不转文字
s2hk deck.pptx -o out.pptx --mode none --ea "楷体" --latin "Arial"

# 已是繁体但用字混乱 → 规整为香港用字
s2hk already_trad.pptx --mode s2hk --via-simplified

# 转换后自动质检（结构 / 字体 / 全文残留）
s2hk deck.pptx -o out.pptx --verify

# 先看会改多少，不落盘；或输出 JSON 便于解析
s2hk deck.pptx --dry-run
s2hk deck.pptx -o out.pptx --verify --json
```

## 转换模式 `--mode`

| mode | 方向 | 说明 |
|------|------|------|
| `s2hk`（默认） | 简 → 港繁 | 香港用字：裏面、啓用、覆盤、賬户、衞生（保留「户」不用「戶」） |
| `s2t` | 简 → 繁 | OpenCC 标准繁体 |
| `s2tw` / `s2twp` | 简 → 台繁 | 台湾正体；`s2twp` 额外做台湾用词替换 |
| `t2s` / `hk2s` / `tw2s` / `tw2sp` | 繁 → 简 | 反向转换 |
| `none` | 不转文字 | 仅改字体，必须配合 `--ea` / `--latin` |

香港 vs 台湾：香港用 **裏**面、**啓**用、賬**户**、**衞**生；台湾用 **裡**面、**啟**用、帳**戶**、**衛**生。面向香港政府 / 公众场景用 `s2hk`。

`--via-simplified`：输入**已经是繁体**但字形混杂时使用，先 `t2s` 归一回简体再转目标 mode，干净地统一为港式（`戶→户`、`纔→才`、`裡→裏`、`啟→啓`）。

## 字体 `--ea` / `--latin` / `--font-scope`

- `--ea` 中文（East-Asian）字体名，`--latin` 英文字体名（同时写入 `<a:cs>`）。
- 字体名要用用户说的那个字面值（「楷体」是最稳的跨平台写法）。
- `--font-scope` 默认 `slides`：**只改正文页**。改母版/版式会波及整份模板观感，除非用户明确要求，否则不要用 `--font-scope all`。
- `font props set: 0` **不一定是失败**：从未设过字符格式的 run 没有 `<a:rPr>` 可写入，它继承主题字体，属正常。

## PPTX 处理范围 `--scope`

默认 `all`。**不要**随意缩小，否则会出现「正文页是繁体、换个版式又冒出简体」。

| scope | 部件 | 为什么重要 |
|---|---|---|
| `slides` | `ppt/slides/` | 正文页 |
| `layouts` | `ppt/slideLayouts/` | 版式上的标题/占位文字，换版式就露出来 |
| `masters` | `slideMasters` / `notesMasters` / `handoutMasters` | 母版级文字 |
| `notes` | `ppt/notesSlides/` | 讲者备注 |
| `diagrams` | `ppt/diagrams/` | **SmartArt**：文字同时存在于 `data*.xml`，只改幻灯片缓存的话，用户在 PowerPoint 里一编辑该图形就会重新渲染、变回简体 |
| `charts` | `ppt/charts/` | 图表标题与轴标签 |
| `props` | `docProps/core.xml` | 文件属性里的标题/主题 |

用法：`--scope all`（默认）/ `--scope slides` / `--scope slides,notes`。

**不会被改动的**：字体名（`腾讯体`/`宋体`）、形状名、主题名等属性值 —— 这些是标识符，改了会破坏文件。

## 验证转换是否彻底

**不要**用「固定简体字表」查残留 —— 香港繁体会合理保留一些看起来像简体的字（`户`、`峰`、`兌`），固定字表必然误报。

正确做法是**二次转换法**：对输出再跑一遍同一 mode，若 `convert(文本) == 文本` 对所有文本成立，即已彻底转换。直接用内置质检即可：

```bash
s2hk deck.pptx -o out.pptx --verify
```

输出包含：页数、文字 run 数、字体分布是否一致、媒体文件是否字节级一致、zip 条目顺序、**全文残留数**（含版式/母版/备注，不只正文页）。

注意：**改了字体时 `font profile same: False` 是预期结果**，不是失败；只有在本次未传 `--ea/--latin` 却出现 False 才需要排查。

偶发误报：OpenCC 的词组规则在「把已转好的繁体再喂回去」时可能想改成更冷僻的异体（如 `適合→適閤`、`團伙→團夥`）。判断方法是拿**简体原词**跑一遍 —— `s2hk('适合')='適合'`，所以第一遍的结果才是标准，保持不动。

## 退出码

| code | 含义 |
|---|---|
| 0 | 成功 |
| 1 | 依赖缺失或文件不存在 |
| 2 | 输入不受支持（`.docx` 等），或 `.pptx` 含路径穿越条目被拒 |
| 3 | `--verify` 未通过 |

## 支持的输入

| 形式 | 例子 | 支持 |
|---|---|---|
| 内联文本 | `--text "用户启用..."` | ✅ |
| UTF-8 纯文本 | `.txt` `.md` `.csv` `.srt` | ✅ |
| PPT | `.pptx` | ✅ 文字 + 可选字体 |
| Office / 排版类二进制 | `.docx` `.xlsx` `.pdf` `.key` `.pages` | ❌ exit 2 |
| 非 UTF-8 文本（GBK / Big5） | — | ❌ 需先另存为 UTF-8 |

「纯文本文件」= 用记事本 / VS Code 打开就能看到原文的文件。`.docx` 看着像文档，实为 zip 二进制容器，**不属于纯文本**。

## 已知边界

- PPTX 中**图片 / 截图 / 视频封面里的文字**属图像内容，无法编辑，保持原状。
- **内嵌 OLE 对象**（如嵌入的 Excel 表）内部文字不在处理范围。
- OpenCC 只做**字形**转换，不改用词习惯（「程序」不会自动变成港式「程式」）。用词本地化需另行处理。
- 若某段落转换后字数变化（`s2twp` 等用词替换会），该段回退为逐 run 转换并提示 `fell back to per-run conversion`。

## 安全

- `.pptx` 属不可信输入：解压前校验每个 zip 条目，拒绝绝对路径与 `../` 路径穿越（Zip Slip），以 exit 2 报错。
- 输出始终写入**新文件**，不覆盖源文件。

## 变更说明（v2）

旧 skill `pptx-s2t-font-convert` 的字体能力已并入本 skill，成为 `--ea` / `--latin` 参数；`verify.py` 升级为内置 `--verify`。若同时看到两个 skill，说明装了旧版本，重跑 `install.sh` 会自动移除旧的。
