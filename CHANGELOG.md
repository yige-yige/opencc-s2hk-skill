# 更新日志

本项目遵循[语义化版本](https://semver.org/lang/zh-CN/)。

## v2.0.0 — 2026-10-22

### 破坏性变更

- **`pptx-s2t-font-convert` 已合并进 `opencc-s2hk`，不再单独存在。**
  两个 skill 原本都基于 OpenCC、都能处理 `.pptx`，差别只有"要不要改字体"——
  这不是用户心智里的分界，结果是两份 SKILL.md 各花十行说明"什么时候该用另一个"，
  agent 还经常选错。现在字体能力变成主脚本的可选参数：

  ```bash
  # v1
  convert.py in.pptx out.pptx --ea 楷体 --latin Arial --mode s2hk
  # v2
  s2hk in.pptx -o out.pptx --ea 楷体 --latin Arial
  ```

  **迁移**：重跑 `install.sh` 会自动移除旧 skill（备份到
  `~/.local/share/opencc-s2hk-skill/backups/`）。全局命令 `s2hk-pptx-font`
  保留为薄包装并打印废弃提示，老的调用方式仍可工作。

- `verify.py` 下线，升级为内置的 `--verify`。旧版是 3 个位置参数、且
  `font identical: False` 容易被误读成失败；新版知道本次是否要改字体，
  只在"不该变却变了"时才判失败。

### 新增

- `--verify`：转换后质检结构 / 字体分布 / 媒体字节 / zip 条目顺序 /
  **全文残留**（含版式、母版、备注，不只正文页）。
- `--dry-run`：只报告会改多少，不落盘。
- `--json`：机器可读输出，便于 agent 解析而不是读 stdout 文本。
- `--mode none`：纯字体处理，不转文字。
- `--font-scope`：字体作用范围，默认 `slides`（改母版会波及整份模板观感）。
- 规范化退出码：0 成功 / 1 依赖或 IO / 2 输入不支持或被安全拒绝 / 3 质检未通过。

### 改进

- `description` 从 709 字符压到 196。Codex 的技能列表只占模型上下文的 2%，
  过长会被截断，反而降低命中率。

## v1.1.0 — 2026-10-22

### 修复

- **处理范围过窄**：旧实现只扫 `ppt/slides/`，在一份 5 页 PPT 上实测漏转
  74 处正文文字（版式 23 个部件、母版 3 个、备注 1 个）。新增 `--scope`，
  默认 `all`，覆盖 slides/layouts/masters/notes/diagrams/charts/props。
  SmartArt 尤其关键：文字也存在 `diagrams/data*.xml`，只改幻灯片缓存的话，
  用户在 PowerPoint 里编辑该图形就会重新渲染、转换被打回原形。残留 74 → 0。
- **Zip Slip**：`.pptx` 属不可信输入，旧实现直接 `extractall`。现在解压前
  校验每个条目，拒绝绝对路径与 `../` 穿越。
- **跨 run 词组被拆导致误转**：PPT 常把一个词拆进多个 `<a:t>`（实测某文件
  41% 段落被拆）。逐节点转换让 OpenCC 词组规则失效：`头发 → 頭發`（错，
  应为 `頭髮`）。改为按段落聚合转换再按原长度回填。
- **zip 重打包不保真**：旧实现用 `ZIP_DEFLATED` 统一重压，把原本 STORED 的
  19 个条目重新压缩。现逐条目沿用原压缩方式 / 顺序 / 时间戳。

### 工程化

- 抽出 `scripts/pptx_core.py` 共享核心，CI 校验各 skill 副本一致。
- 新增 `tests/test_convert.py`：现场生成 fixture 的端到端回归测试。

## v1.0.1 — 2026-10-21

- 修复备份目录留在 skills 目录内，被 agent 当成同名技能重复加载。
- 修复 bash 把 `$var` 后紧跟的中文字符并入变量名，导致 `unbound variable`
  中断安装；并在 CI 中加入该类写法的静态检查。

## v1.0.0 — 2026-10-21

首个版本：`opencc-s2hk` 与 `pptx-s2t-font-convert` 两个 skill，
三层覆盖安装器（标准共享目录 + 各家专属目录 + 全局命令）。
