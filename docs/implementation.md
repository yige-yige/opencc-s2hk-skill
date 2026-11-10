# 实现与测试

## 处理原理

- **段落级聚合转换**：PPT 会把一个词拆成多个 `<a:t>`。先合并整段再转换，可避免「头」+「发」被错转为「頭發」而非「頭髮」。
- **完整 PPTX 范围**：除正文页外，同时处理版式、母版、备注、SmartArt 数据源、图表和文件属性。
- **保真重打包**：沿用原 ZIP 条目顺序、压缩方式与时间戳；图片、图标和嵌入对象字节不变。
- **原子输出**：先写入目标同目录下的临时文件，成功后再原子替换目标。

## 安全措施

- 解压 PPTX 前检查 ZIP 条目，拒绝绝对路径和 `../` 路径穿越。
- 输入与输出指向同一文件时立即拒绝，包括符号链接和硬链接别名。
- 安装器备份已有的同名命令；卸载器只删除有本项目所有权记录的命令和 XDG 链接。

## 仓库结构

```text
skills/opencc-s2hk/
├── SKILL.md
├── references/advanced.md
└── scripts/
    ├── s2hk_convert.py
    └── pptx_core.py
docs/
├── installation.md
├── cli.md
├── implementation.md
└── troubleshooting.md
tests/
├── test_convert.py
├── test_safety.py
└── test_installer_safety.py
install.sh
uninstall.sh
AGENT_SNIPPET.md
```

## 本地校验

```bash
python3 .github/validate_skills.py
python3 tests/test_convert.py
python3 tests/test_safety.py
python3 tests/test_installer_safety.py
bash -n install.sh
bash -n uninstall.sh
```

`test_convert.py` 需要 `opencc`、`lxml` 和 `python-pptx`；安全回归测试不需要第三方依赖。

GitHub Actions 配置目前位于 `ci/validate.yml`。启用方法见文件头部说明。
