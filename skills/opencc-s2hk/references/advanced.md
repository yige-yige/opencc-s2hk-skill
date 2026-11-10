# opencc-s2hk 进阶指南

仅当任务涉及非默认转换、PPTX 字体 / scope、质检或依赖故障时读取本文件。

## 依赖与调用

`install.sh` 会记录可用的 Python 解释器。手动安装时所需依赖：

```bash
python3 -m pip install opencc
python3 -m pip install lxml python-pptx
```

`lxml` 仅 PPTX 需要；`python-pptx` 仅 `--verify` 需要。不要因依赖缺失改用正则或手写字表。若系统 Python 受 PEP 668 限制，优先使用仓库安装器创建的独立 venv。

## 转换模式

| mode | 方向 | 说明 |
|---|---|---|
| `s2hk` | 简 → 港繁 | 默认；香港字形 |
| `s2t` | 简 → 繁 | OpenCC 通用繁体 |
| `s2tw` / `s2twp` | 简 → 台繁 | `s2twp` 额外转换台湾用词 |
| `t2s` / `hk2s` / `tw2s` / `tw2sp` | 繁 → 简 | 反向转换 |
| `none` | 不转文字 | 仅改 PPTX 字体，必须配合 `--ea` / `--latin` |

`--via-simplified` 适用于已是繁体但字形混杂的输入：先经 `t2s` 归一，再转到目标 mode。

## PPTX 文字范围

`--scope all` 是默认值，包含：

| scope | 部件 |
|---|---|
| `slides` | 正文页 |
| `layouts` | 版式 |
| `masters` | 幻灯片 / 备注 / 讲义母版 |
| `notes` | 讲者备注 |
| `diagrams` | SmartArt 数据源 |
| `charts` | 图表文字 |
| `props` | 文件属性 |

只在用户明确要求时使用 `--scope slides` 或逗号组合。SmartArt 文字也保存在 `ppt/diagrams/data*.xml`；漏掉它会导致 PowerPoint 重新渲染后变回原文。

## PPTX 字体

- `--ea` 设置 East-Asian 字体；`--latin` 设置拉丁字体，同时写入 complex-script 字体。
- `--font-scope` 默认为 `slides`。除非用户要求改模板，不要使用 `--font-scope all`。
- `font props set: 0` 不一定是失败；没有显式字符格式的 run 会继承主题字体。
- 字体名只是写入 PPTX 的字面值；目标电脑没有该字体时，PowerPoint 可能替换字体。

## 质检

```bash
s2hk deck.pptx -o out.pptx --verify
```

`--verify` 检查页数、文字 run、字体分布、媒体字节、ZIP 条目顺序与全文残留。改字体时 `font profile same: False` 是预期结果。

残留检查使用二次转换法，不用固定简体字表。OpenCC 偶尔会在二次转换时倾向冷僻异体；遇到此类差异时，用简体原词单独转换来判定标准结果。

## 输入与退出码

| code | 含义 |
|---|---|
| 0 | 成功 |
| 1 | 依赖缺失或文件不存在 |
| 2 | 不支持的输入、输入输出同文件，或拒绝恶意 PPTX |
| 3 | `--verify` 未通过 |

纯文本必须是 UTF-8。`.docx` / `.xlsx` / `.pdf` / `.key` / `.pages` 不在支持范围。

## 安全不变量

- PPTX 解压前拒绝绝对路径与 `../` 路径穿越。
- 输入与输出不得指向同一文件，包括符号链接和硬链接别名。
- 输出先写入同目录临时文件，成功后再原子替换目标。
