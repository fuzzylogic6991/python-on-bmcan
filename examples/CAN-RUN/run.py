#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
一键运行 CAN 诊断测试工程
- 根据配置自动选择 32/64 位 Python
- 设置 PYTHONPATH 并运行 uds.py
"""

import os
import sys
import subprocess

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PARENT = os.path.normpath(os.path.join(SCRIPT_DIR, "..", ".."))

# ── 项目配置（切换项目改这里） ──
CAN_PROJECT = "config_duola"

# config_duola 使用 32 位安全算法 DLL，需要 32 位 Python
if CAN_PROJECT == "config_duola":
    PYTHON_EXE = os.path.join(
        os.environ["LOCALAPPDATA"],
        r"Programs\Python\Python311-32\python.exe"
    )
else:
    PYTHON_EXE = os.path.join(
        os.environ["LOCALAPPDATA"],
        r"Programs\Python\Python312\python.exe"
    )


def main():
    # 找到要运行的 .py 文件（排除配置文件和工具脚本）
    target_py = None
    for f in sorted(os.listdir(SCRIPT_DIR)):
        if f.endswith(".py") and not f.startswith("config_") and f not in ("run.py", "setup.py"):
            target_py = os.path.join(SCRIPT_DIR, f)
            break

    if target_py is None:
        print("[ERROR] No runnable .py file found")
        input("Press Enter to exit...")
        sys.exit(1)

    # 检查 Python 是否存在
    if not os.path.exists(PYTHON_EXE):
        print(f"[ERROR] Python not found: {PYTHON_EXE}")
        print("  Please run setup.bat first.")
        input("Press Enter to exit...")
        sys.exit(1)

    print("=" * 60)
    print(f"  Running: {os.path.basename(target_py)}")
    print(f"  Config: {CAN_PROJECT}")
    print(f"  Python: {PYTHON_EXE}")
    print(f"  Dir: {SCRIPT_DIR}")
    print("=" * 60)
    print()

    # 设置环境变量并运行
    env = os.environ.copy()
    env["CAN_PROJECT"] = CAN_PROJECT
    env["PYTHONPATH"] = PARENT + ";" + env.get("PYTHONPATH", "")

    proc = subprocess.run(
        [PYTHON_EXE, target_py],
        cwd=SCRIPT_DIR,
        env=env,
    )

    print()
    sys.exit(proc.returncode)


if __name__ == "__main__":
    main()
