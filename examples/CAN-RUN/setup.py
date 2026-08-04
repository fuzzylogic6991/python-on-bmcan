#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
CAN-RUN 一键环境安装脚本
- 自动检测/安装 32 位 Python 3.11（如需要）
- 安装 python-can 及所有依赖
- 32 位 Python 自动使用 pandas<2.1 + numpy<2（兼容 win32）
"""

import os
import sys
import subprocess
import urllib.request
import tempfile

# ── 配置 ──
PYTHON_32_DIR = os.path.join(os.environ["LOCALAPPDATA"], r"Programs\Python\Python311-32")
PYTHON_32_EXE = os.path.join(PYTHON_32_DIR, "python.exe")
PYTHON_VER = "3.11.9"
PYTHON_URL = f"https://www.python.org/ftp/python/{PYTHON_VER}/python-{PYTHON_VER}.exe"

# 项目根目录（脚本在 examples/CAN-RUN/ 下）
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.normpath(os.path.join(SCRIPT_DIR, "..", ".."))


def run(cmd, **kwargs):
    """执行命令，返回 (returncode, output)"""
    result = subprocess.run(cmd, capture_output=True, text=True, **kwargs)
    return result.returncode, result.stdout + result.stderr


def python_bits(python_exe):
    """检测 Python 位数"""
    try:
        proc = subprocess.run(
            [python_exe, "-c", "import sys; print(64 if sys.maxsize > 2**32 else 32)"],
            capture_output=True, text=True, timeout=10,
        )
        return int(proc.stdout.strip())
    except Exception:
        return None


def install_32bit_python():
    """下载并安装 32 位 Python 3.11"""
    print("[0/6] Checking 32-bit Python (Python311-32) ...")
    if os.path.exists(PYTHON_32_EXE):
        bits = python_bits(PYTHON_32_EXE)
        if bits == 32:
            print("       [OK] 32-bit Python ready")
            return True
    print(f"       Not found. Downloading Python {PYTHON_VER} 32-bit ...")
    installer_path = os.path.join(tempfile.gettempdir(), f"python-{PYTHON_VER}-x86.exe")
    try:
        urllib.request.urlretrieve(PYTHON_URL, installer_path)
    except Exception as e:
        print(f"[ERROR] Failed to download Python installer: {e}")
        return False
    if not os.path.exists(installer_path):
        print(f"[ERROR] Download failed: {installer_path} not found")
        return False
    print(f"       Installing to {PYTHON_32_DIR} ...")
    print("       (this may take a minute, please wait)")
    rc, _ = run([
        installer_path, "/quiet",
        "InstallAllUsers=0",
        f"TargetDir={PYTHON_32_DIR}",
        "Include_test=0",
        "Include_pip=1",
    ])
    try:
        os.remove(installer_path)
    except Exception:
        pass
    if rc != 0:
        print(f"[ERROR] Python 32-bit installation failed (exit code: {rc})")
        print(f"        You can manually install from: {PYTHON_URL}")
        return False
    if not os.path.exists(PYTHON_32_EXE):
        print("[ERROR] Python 32-bit installed but executable not found")
        return False
    print(f"[OK] Python {PYTHON_VER} 32-bit installed")
    return True


def fail(msg):
    """Print error and exit."""
    print(f"\n[ERROR] {msg}")
    sys.exit(1)


def main():
    print("=" * 60)
    print("  CAN-RUN Environment Setup")
    print("=" * 60)
    print()

    # ── 确定使用哪个 Python ──
    # 优先用已有的 Python 运行本脚本，否则尝试 32 位 Python
    current_bits = 64 if sys.maxsize > 2**32 else 32
    if current_bits == 32:
        # 当前就是用 32 位 Python 运行的，直接用
        python_exe = sys.executable
    elif os.path.exists(PYTHON_32_EXE):
        # 当前是 64 位，但存在 32 位 Python，切换到 32 位
        python_exe = PYTHON_32_EXE
    else:
        # 需要安装 32 位 Python
        python_exe = PYTHON_32_EXE

    # ── Step 0: 安装 32 位 Python（如需要） ──
    if not os.path.exists(PYTHON_32_EXE):
        if not install_32bit_python():
            fail("Cannot proceed without 32-bit Python.")
    else:
        print(f"[0/6] Checking 32-bit Python (Python311-32) ...")
        print("       [OK] 32-bit Python ready")
    print()

    # ── Step 1: 检测位数 ──
    bits = python_bits(python_exe)
    if bits is None:
        fail("Cannot detect Python architecture")
    print(f"[1/6] Detecting Python architecture ...")
    print(f"       Python is {bits}-bit")
    print()

    # ── Step 2: 升级 pip ──
    print("[2/6] Upgrading pip ...")
    rc, out = run([python_exe, "-m", "pip", "install", "--upgrade", "pip", "--quiet"])
    if rc != 0:
        print(f"[WARN] pip upgrade had issues (non-fatal): {out[-200:]}")
    print("[OK] pip ready")
    print()

    # ── Step 3: 安装 python-can ──
    print("[3/6] Installing custom python-can (with bmcan) ...")
    print("       Uninstalling PyPI version if exists ...")
    run([python_exe, "-m", "pip", "uninstall", "python-can", "-y"], timeout=60)
    rc, out = run([python_exe, "-m", "pip", "install", "-e", ROOT, "--quiet"], timeout=120)
    if rc != 0:
        fail(f"python-can install failed:\n{out[-500:]}")
    print("[OK] python-can installed")
    print()

    # ── Step 4: 安装依赖 ──
    print("[4/6] Installing dependencies (pandas / openpyxl / pyserial) ...")
    if bits == 32:
        print("       32-bit Python: using pandas<2.1 + numpy<2 (for compatibility)")
        rc, out = run([python_exe, "-m", "pip", "install",
                        "pandas>=2.0,<2.1", "numpy<2",
                        "openpyxl>=3.0.0", "pyserial>=3.5",
                        "typing_extensions>=3.10.0.0",
                        "--quiet"], timeout=300)
    else:
        req_file = os.path.join(SCRIPT_DIR, "requirements.txt")
        rc, out = run([python_exe, "-m", "pip", "install", "-r", req_file, "--quiet"], timeout=300)
    if rc != 0:
        fail(f"Dependencies install failed:\n{out[-500:]}")
    print("[OK] Dependencies installed")
    print()

    # ── Step 5: 验证 bmcan ──
    print("[5/6] Verifying bmcan interface ...")
    rc, out = run([python_exe, "-c",
                    "import can; ok='bmcan' in can.VALID_INTERFACES; "
                    "print('[OK] bmcan available' if ok else '[FAIL] bmcan not available'); "
                    "exit(0 if ok else 1)"])
    if rc != 0:
        print(out)
        fail("bmcan interface not registered.\n"
             "  Make sure PyPI python-can is not installed.")

    print()
    print("=" * 60)
    print("  [OK] Setup complete! Run run.bat to start.")
    print("=" * 60)


# ── 运行 ──
if __name__ == "__main__":
    main()
    try:
        input("Press Enter to exit...")
    except (EOFError, KeyboardInterrupt):
        pass
