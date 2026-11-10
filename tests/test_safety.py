#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""无第三方依赖的输出安全回归测试。"""
import importlib.util
import os
import shutil
import sys
import tempfile
import zipfile


ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRIPTS = os.path.join(ROOT, "skills", "opencc-s2hk", "scripts")
sys.path.insert(0, SCRIPTS)

spec = importlib.util.spec_from_file_location(
    "s2hk_convert", os.path.join(SCRIPTS, "s2hk_convert.py"))
convert_mod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(convert_mod)

from pptx_core import ensure_distinct_paths, repack  # noqa: E402


def expect_value_error(fn, label):
    try:
        fn()
    except ValueError:
        print("  [pass] %s" % label)
        return
    raise AssertionError(label)


def main():
    tmp = tempfile.mkdtemp(prefix="s2hk-safety-")
    try:
        text_path = os.path.join(tmp, "source.md")
        with open(text_path, "w", encoding="utf-8") as f:
            f.write("源文件\n")
        with open(text_path, "rb") as f:
            original = f.read()
        expect_value_error(
            lambda: convert_mod.convert_text_file(text_path, text_path, lambda s: s),
            "文本输入输出同路径被拒绝")
        with open(text_path, "rb") as f:
            assert f.read() == original
        print("  [pass] 文本源文件保持不变")

        text_out = os.path.join(tmp, "out.md")
        convert_mod.convert_text_file(text_path, text_out, lambda s: s + "已转换\n")
        with open(text_out, encoding="utf-8") as f:
            assert f.read().endswith("已转换\n")
        assert not [n for n in os.listdir(tmp) if n.startswith(".s2hk-")]
        print("  [pass] 文本原子输出且无临时文件残留")

        alias = os.path.join(tmp, "alias.md")
        os.symlink(text_path, alias)
        expect_value_error(
            lambda: ensure_distinct_paths(text_path, alias),
            "指向源文件的符号链接被拒绝")

        src_zip = os.path.join(tmp, "source.pptx")
        extracted = os.path.join(tmp, "parts")
        os.mkdir(extracted)
        with zipfile.ZipFile(src_zip, "w") as zf:
            zf.writestr("ppt/slides/slide1.xml", b"<root/>")
        with zipfile.ZipFile(src_zip) as zf:
            infos = list(zf.infolist())
            zf.extractall(extracted)
        with open(src_zip, "rb") as f:
            zip_original = f.read()
        expect_value_error(
            lambda: repack(src_zip, src_zip, extracted, infos),
            "PPTX 输入输出同路径被拒绝")
        with open(src_zip, "rb") as f:
            assert f.read() == zip_original
        print("  [pass] PPTX 源文件保持不变")

        out_zip = os.path.join(tmp, "out.pptx")
        repack(src_zip, out_zip, extracted, infos)
        assert zipfile.is_zipfile(out_zip)
        assert not [n for n in os.listdir(tmp) if n.startswith(".s2hk-")]
        print("  [pass] PPTX 原子输出且无临时文件残留")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    print("全部通过")
    return 0


if __name__ == "__main__":
    sys.exit(main())
