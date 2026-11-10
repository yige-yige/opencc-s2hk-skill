---
name: opencc-s2hk
description: Convert inline text, UTF-8 text files, and PPTX content between Simplified and Traditional Chinese, defaulting to Hong Kong Traditional (s2hk). Use for 简体转繁体、香港繁体、港式用字规整、PPT 转繁体，或 PPTX 中英文字体统一。 Do not use for DOCX, XLSX, PDF, or font changes outside PPTX.
license: MIT
compatibility: Requires Python 3 and OpenCC; PPTX support also requires lxml, and --verify requires python-pptx. The bundled Bash installer supports macOS and Linux.
metadata:
  description_zh: 简繁转换与 PPTX 字体统一（默认香港繁体）
  description_en: Simplified/Traditional converter with PPTX font control
---

# opencc-s2hk

使用 OpenCC 转换中文字形，默认输出香港繁体 `s2hk`。处理 PPTX 时只修改文字与用户明确指定的字体，保留版式、配色、图片与表格。

## 执行流程

1. 确认输入是内联文本、UTF-8 纯文本文件或 `.pptx`。对 `.docx`、`.xlsx`、`.pdf` 等不支持格式，说明边界并停止。
2. 默认使用 `s2hk`。只有用户明确要求其他繁体标准或反向转换时，才选择其他 `--mode`。
3. 优先调用已安装的 `s2hk`。若不存在，用当前 skill 的绝对路径执行 `python3 <skill-dir>/scripts/s2hk_convert.py`。
4. 输出到新文件；不得把 `-o` 指向输入文件。
5. 任务需要严格质检时，对 PPTX 加 `--verify`，并如实报告 PASS / FAIL。

## 常用命令

```bash
s2hk --text "用户启用多租户"
s2hk notes.md
s2hk deck.pptx
s2hk deck.pptx -o out.pptx --ea "楷体" --latin "Arial"
s2hk old_trad.pptx --via-simplified
s2hk deck.pptx -o out.pptx --verify
```

## 关键规则

- 不要用正则、手写字表或模型自行替换来代替 OpenCC。
- PPTX 文字转换默认使用 `--scope all`，覆盖正文、版式、母版、备注、SmartArt、图表和文件属性。除非用户明确要求，不要缩小范围。
- 默认不改字体。仅在用户要求时传入 `--ea` / `--latin`；字体默认只应用于正文页。
- 输入已是繁体但港台字形混杂时，使用 `--via-simplified` 先归一再转目标模式。
- 香港繁体会合理保留 `户`、`峰`、`兌` 等看似简体的字。不要用固定简体字表判定残留。

## 边界

- 图片、截图、视频封面与内嵌 OLE 对象中的文字不会被转换。
- OpenCC 只转换字形，不负责港台词汇本地化。
- 非 UTF-8 纯文本需先转码。

## 按需读取

当任务涉及非默认模式、PPTX scope、字体、质检、退出码或依赖排错时，读取 [references/advanced.md](references/advanced.md)。
