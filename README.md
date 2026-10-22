# opencc-s2hk · 简体 → 香港繁体 Agent Skill

把 PPT / Markdown / 文本里的简体中文转成**香港繁体（s2hk 港式用字）**，版式、字体、图片、表格、配色全部保持原样。

基于 [Agent Skills 开放标准](https://agentskills.io/specification)，**一份 SKILL.md 跨 agent 通用**：CodeBuddy、WorkBuddy、Codex、Claude Code、Cursor、Gemini CLI、GitHub Copilot、Cline、Zed、Warp、Amp、OpenCode、Windsurf、Trae、通义灵码…… 以及任何能执行终端命令的 agent。

| Skill | 用途 |
|---|---|
| `opencc-s2hk` | 内联文本 / `.txt` `.md` `.csv` `.srt` / `.pptx` 文字转换 |
| `pptx-s2t-font-convert` | `.pptx` 转换 **+ 中英文字体统一 +** 质检 |

分工：输入不是 `.pptx`、或不涉及字体 → `opencc-s2hk`；`.pptx` 且要改字体 → `pptx-s2t-font-convert`。两边 SKILL.md 已写好互相转交规则，agent 会自己选。

---

## 安装

### 方式 A：官方 skills CLI（一条命令，覆盖 78 个 agent）

需要 Node.js ≥ 22.20。

```bash
npx skills add yige-yige/opencc-s2hk-skill --list    # 先看有哪些 skill、会装到哪些 agent
npx skills add yige-yige/opencc-s2hk-skill --all     # 全部 skill 装到全部已检测到的 agent
```

常用变体：

```bash
npx skills add yige-yige/opencc-s2hk-skill -g                          # 全局安装（默认是只装当前项目！）
npx skills add yige-yige/opencc-s2hk-skill -a codex -a claude-code -g  # 只装指定 agent
npx skills add yige-yige/opencc-s2hk-skill --skill opencc-s2hk -g      # 只装核心 skill
npx skills add yige-yige/opencc-s2hk-skill --all --copy                # 复制而非软链接
npx skills update                                                      # 日后更新
```

装完后**首次使用前需要装 Python 依赖**（CLI 只负责分发文件）：

```bash
python3 -m pip install --user opencc lxml python-pptx
```

> ⚠️ 该 CLI 的 agent 注册表里**有 CodeBuddy、没有 WorkBuddy**。要装进 WorkBuddy，请用方式 B。

### 方式 B：仓库自带安装器（零 Node 依赖，自动处理 Python 环境）

```bash
git clone https://github.com/yige-yige/opencc-s2hk-skill.git
bash opencc-s2hk-skill/install.sh --list    # 只检测，不写任何文件
bash opencc-s2hk-skill/install.sh           # 正式安装
```

它比方式 A 多做三件事：

1. **自动装 Python 依赖**：`pip` → `--user` → `--break-system-packages` → 独立 venv 四级兜底；`opencc` 无预编译包时自动换 `opencc-python-reimplemented`。把最终可用的解释器写进 `scripts/.python`，脚本检测到依赖缺失会自动切过去 —— agent 用哪个 `python3` 调用都不会 `ModuleNotFoundError`。
2. **覆盖 WorkBuddy** 等注册表里没有的 agent。
3. **生成全局命令** `~/.local/bin/s2hk`，给完全不支持 SKILL.md 的 agent 兜底。

| 参数 | 作用 |
|---|---|
| `--list` | 只检测本机装了哪些 agent，不写文件 |
| `--all` | 连本机还没装的 agent 目录也一起建好 |
| `--link` | 用软链接代替复制，升级只改共享目录一处 |
| `--core-only` | 只装 `opencc-s2hk` |
| `--no-deps` | 跳过 Python 依赖安装 |
| `--add-path` | 把 `~/.local/bin` 写进 `~/.zshrc` |

卸载：`bash opencc-s2hk-skill/uninstall.sh`

### 方式 C：手动复制

```bash
mkdir -p ~/.agents/skills && cp -R skills/* ~/.agents/skills/
```

`~/.agents/skills/` 是 Agent Skills 标准的共享位置，Codex / Cursor / Gemini CLI / Copilot / Cline / Zed / Warp / Amp / OpenCode 等都读它。Claude Code（`~/.claude/skills/`）、CodeBuddy（`~/.codebuddy/skills/`）、WorkBuddy（`~/.workbuddy/skills/`）用自己的目录，需另外复制一份。

---

## 怎么用

重启 / 新开 agent 会话后用大白话说即可：

| 你说 | 会发生什么 |
|---|---|
| 「把这个 PPT 转成香港繁体」（拖入 .pptx） | 输出 `xxx.s2hk.pptx`，只改文字 |
| 「这份 Markdown 转繁体」 | 输出 `xxx.s2hk.md` |
| 「这段话转成港式用字：用户启用多租户」 | 直接回繁体文本 |
| 「PPT 转繁体，同时中文用楷体、英文用 Arial」 | 走字体增强版，一次搞定 |
| 「这份文件已是繁体但 戶/纔/裡 用字混乱，统一成香港写法」 | 走 `--via-simplified` 规整 |

Codex 里可显式点名 `$opencc-s2hk`，用 `/skills` 查看是否已加载。

命令行直接用（方式 B 安装后）：

```bash
s2hk --text "用户启用多租户"        # 内联文本
s2hk notes.md                       # → notes.s2hk.md
s2hk deck.pptx                      # → deck.s2hk.pptx
s2hk deck.pptx -o out.pptx --mode s2hk
s2hk old_trad.pptx --via-simplified # 已是繁体但用字混乱 → 规整为港式
s2hk-pptx-font in.pptx out.pptx --ea "楷体" --latin "Arial" --mode s2hk
```

转换模式：`s2hk`（默认，香港用字：裏面/啓用/賬户/衞生）、`s2t`（通用繁体）、`s2tw` `s2twp`（台湾正体）、`t2s` `hk2s` `tw2s`（反向转简体）。

### PPTX 处理范围

默认转换**全部**含文字的部件，不只是正文页：

| scope | 部件 | 说明 |
|---|---|---|
| `slides` | 正文页 | |
| `layouts` / `masters` | 版式 / 母版 | 否则换个版式就冒出简体 |
| `notes` | 备注页 | 讲者备注 |
| `diagrams` | **SmartArt** | 文字也存在 `data*.xml`，只改幻灯片缓存的话，在 PowerPoint 里编辑该图形会重新渲染、变回简体 |
| `charts` / `props` | 图表 / 文件属性 | |

```bash
s2hk deck.pptx                      # 默认 --scope all
s2hk deck.pptx --scope slides       # 只转正文页
s2hk deck.pptx --scope slides,notes # 自选组合
```

字体名、形状名等属性值**不会**被改动（它们是标识符，改了会坏文件）。

---

## 实现要点

几个容易写错、这里已经处理掉的地方：

- **段落级聚合转换**：PPT 常把一个词拆进多个 `<a:t>`（设计师加粗半个词就会这样）。逐节点转换会让 OpenCC 的词组规则失效 —— `头发` 拆成 `头`+`发` 会得到 `頭發`（错），正确是 `頭髮`。这里按段落拼接后转换再按原长度回填。
- **Zip Slip 防护**：`.pptx` 是不可信输入，解压前校验每个 zip 条目，拒绝绝对路径与 `../` 穿越，exit 2 报错。
- **保真重打包**：逐条目沿用原压缩方式/顺序/时间戳。直接 `ZIP_DEFLATED` 重打包会把原本 STORED 的条目重新压缩，对已压缩的 PNG/JPEG 既不省空间又损失保真度。
- **媒体零改动**：图片/图标/嵌入对象字节级原样搬运。

---

## 仓库结构

```
.
├── skills/                         # ← 官方 CLI 扫描这里（skills/<name>/SKILL.md）
│   ├── opencc-s2hk/
│   │   ├── SKILL.md
│   │   └── scripts/{s2hk_convert.py,pptx_core.py}
│   └── pptx-s2t-font-convert/
│       ├── SKILL.md
│       └── scripts/{convert.py,verify.py,pptx_core.py}
├── install.sh                      # 通用安装器（三层覆盖）
├── uninstall.sh
├── AGENT_SNIPPET.md                # 给不支持 SKILL.md 的 agent 贴的规则片段
├── tests/test_convert.py           # 端到端回归测试（22 项断言）
├── .github/
│   └── validate_skills.py          # 自检：frontmatter / 脚本语法 / 转换幂等性
├── ci/validate.yml                 # GitHub Actions（待启用，见文件头注释）
├── LICENSE
└── README.md
```

`pptx_core.py` 是两个 skill 共用的处理核心（安全解压 / 处理范围 / 段落级转换 / 保真重打包）。skill 必须自包含不能跨目录 import，所以各存一份，CI 会校验两份内容一致。

本地跑测试与自检：

```bash
python3 tests/test_convert.py       # 端到端：跨 run 拆词 / Zip Slip / 保真度 / 幂等性
python3 .github/validate_skills.py  # 静态检查：frontmatter / 脚本语法 / shell 陷阱
```

---

## 常见问题

**`command not found: python3`** → `xcode-select --install`（macOS）或 `brew install python`，再重跑安装器。

**`externally-managed-environment`** → 用方式 B 安装，安装器会自动建独立 venv。

**`ModuleNotFoundError: opencc`** → 用方式 A 装的，补一条 `python3 -m pip install --user opencc lxml python-pptx`；或改用方式 B。

**agent 说不知道这个技能** →
1. 确认文件在：`ls ~/.agents/skills ~/.codebuddy/skills ~/.workbuddy/skills`
2. 完全退出客户端再开（Codex 用 `/skills` 看列表）
3. 目录名必须正好是 `opencc-s2hk/SKILL.md`，不能多套一层
4. 注意方式 A **默认装到当前项目**而非全局，换目录就失效 —— 加 `-g`
5. 技能装太多时 Codex 会因上下文预算省略部分技能，可在 `~/.codex/config.toml` 调 `skills.max_context_tokens`（上限 10000）
6. 兜底：把 `AGENT_SNIPPET.md` 贴进该 agent 的规则文件，让它直接用 `s2hk` 命令

**PPT 里有几个字没转** → 默认已覆盖版式/母版/备注/SmartArt。若仍有残留，通常是：① 文字在图片 / 截图里（图像内容改不了）；② 内嵌 OLE 对象（如嵌入的 Excel）里的文字。

**繁体里有些字看着像简体（户、峰、兌）** → 这是香港标准用字，不是漏转。**不要**用固定简体字表查残留，会误报；正确校验法是对输出再跑一遍同一 mode，输出不变即已转彻底。

**「程序」没变成「程式」** → OpenCC 只做字形转换，不做港台用词替换，用词本地化需人工处理。

---

## 致谢与许可

基于 [OpenCC](https://github.com/BYVoid/OpenCC)（Apache-2.0）、`lxml`、`python-pptx`。本仓库代码 MIT，见 [LICENSE](./LICENSE)。
