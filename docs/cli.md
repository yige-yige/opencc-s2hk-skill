# CLI 手册

## 基本用法

```bash
s2hk --text "用户启用多租户"        # 内联文本
s2hk notes.md                       # → notes.s2hk.md
s2hk deck.pptx                      # → deck.s2hk.pptx
s2hk deck.pptx -o out.pptx
s2hk deck.pptx --dry-run            # 只计算，不写文件
s2hk deck.pptx -o out.pptx --json   # 机器可读输出
```

## 转换模式

| mode | 方向 | 说明 |
|---|---|---|
| `s2hk` | 简 → 港繁 | 默认；香港字形 |
| `s2t` | 简 → 繁 | OpenCC 通用繁体 |
| `s2tw` / `s2twp` | 简 → 台繁 | `s2twp` 额外做台湾用词替换 |
| `t2s` / `hk2s` / `tw2s` / `tw2sp` | 繁 → 简 | 反向转换 |
| `none` | 不转文字 | 仅改 PPTX 字体，必须配合 `--ea` / `--latin` |

香港常用 **裏**面、**啓**用、賬**户**、**衞**生；台湾常用 **裡**面、**啟**用、帳**戶**、**衛**生。

输入已是繁体但港台字形混杂时：

```bash
s2hk old_trad.pptx --mode s2hk --via-simplified
```

## PPTX 处理范围

默认 `--scope all`：

| scope | 部件 |
|---|---|
| `slides` | 正文页 |
| `layouts` | 版式 |
| `masters` | 幻灯片 / 备注 / 讲义母版 |
| `notes` | 讲者备注 |
| `diagrams` | SmartArt 数据源 |
| `charts` | 图表文字 |
| `props` | 文件属性 |

```bash
s2hk deck.pptx --scope all
s2hk deck.pptx --scope slides
s2hk deck.pptx --scope slides,notes
```

字体名、形状名、主题名等标识符不会被文字转换。

## PPTX 字体

```bash
s2hk deck.pptx -o out.pptx --ea "楷体" --latin "Arial"
s2hk deck.pptx -o out.pptx --mode none --ea "楷体"
```

- `--ea` 设置 East-Asian 字体；`--latin` 设置拉丁字体。
- `--font-scope` 默认 `slides`。使用 `--font-scope all` 会修改母版与版式的字体设定。
- 字体名只是写入文件的字面值；目标电脑没有该字体时，PowerPoint 可能替换字体。
- `font props set: 0` 不一定是失败；某些 run 只继承主题字体。

## 质检

```bash
s2hk deck.pptx -o out.pptx --verify
s2hk deck.pptx -o out.pptx --verify --json
```

质检覆盖页数、文字 run、字体分布、媒体字节、ZIP 条目顺序与全文残留。当命令明确修改字体时，`font profile same: False` 是预期结果。

香港繁体会保留 `户`、`峰`、`兌` 等看似简体的字，因此不使用固定简体字表检查残留。

## 输入限制与退出码

| code | 含义 |
|---|---|
| 0 | 成功 |
| 1 | 依赖缺失或文件不存在 |
| 2 | 输入不支持、输入输出同文件，或 PPTX 未通过安全检查 |
| 3 | `--verify` 未通过 |

纯文本必须是 UTF-8。`.docx` / `.xlsx` / `.pdf` / `.key` / `.pages` 等二进制文件会被拒绝。输出先写入同目录临时文件，成功后再原子替换目标。
