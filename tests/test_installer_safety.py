#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""无第三方依赖的安装/卸载所有权回归测试。"""
import os
import shutil
import subprocess
import sys
import tempfile


ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def run(script, home, path):
    args = ["bash", os.path.join(ROOT, script)]
    args.append("--no-deps" if script == "install.sh" else "--keep-env")
    return subprocess.run(
        args,
        env=dict(os.environ, HOME=home, PATH=path),
        stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True,
    )


def main():
    tmp = tempfile.mkdtemp(prefix="s2hk-installer-")
    try:
        fake_bin = os.path.join(tmp, "fake-bin")
        os.mkdir(fake_bin)
        fake_python = os.path.join(fake_bin, "python3")
        with open(fake_python, "w", encoding="utf-8") as f:
            f.write("#!/bin/sh\nprintf '%s\\n' '用户啓用多租戶數據隔離機制'\n")
        os.chmod(fake_python, 0o755)
        path = fake_bin + os.pathsep + os.environ.get("PATH", "")

        home = os.path.join(tmp, "owned")
        os.makedirs(os.path.join(home, ".local", "bin"))
        command = os.path.join(home, ".local", "bin", "s2hk")
        with open(command, "w", encoding="utf-8") as f:
            f.write("original user command\n")
        installed = run("install.sh", home, path)
        assert installed.returncode == 0, installed.stdout
        with open(command, encoding="utf-8") as f:
            assert "managed-by: opencc-s2hk-skill" in f.read()
        removed = run("uninstall.sh", home, path)
        assert removed.returncode == 0, removed.stdout
        with open(command, encoding="utf-8") as f:
            assert f.read() == "original user command\n"
        print("  [pass] 同名用户命令在卸载后恢复")

        foreign_home = os.path.join(tmp, "foreign")
        xdg_parent = os.path.join(foreign_home, ".config", "agents")
        foreign_target = os.path.join(foreign_home, "foreign-skills")
        os.makedirs(xdg_parent)
        os.mkdir(foreign_target)
        foreign_link = os.path.join(xdg_parent, "skills")
        os.symlink(foreign_target, foreign_link)
        removed = run("uninstall.sh", foreign_home, path)
        assert removed.returncode == 0, removed.stdout
        assert os.path.islink(foreign_link)
        assert os.readlink(foreign_link) == foreign_target
        print("  [pass] 非本项目创建的 XDG 软链接被保留")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    print("全部通过")
    return 0


if __name__ == "__main__":
    sys.exit(main())
