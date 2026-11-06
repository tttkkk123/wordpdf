"""运行环境探测：本机到底装了哪种能干活的东西。

这个模块是整个插件"选路"的依据 —— 不猜，只实事求是地探测。
"""

from __future__ import annotations

import os
import shutil
import sys
from pathlib import Path
from typing import Optional

IS_WIN = os.name == "nt"

_WINWORD_CANDIDATES = [
    r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE",
    r"C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE",
    r"C:\Program Files\Microsoft Office\Office16\WINWORD.EXE",
    r"C:\Program Files (x86)\Microsoft Office\Office16\WINWORD.EXE",
    r"C:\Program Files\Microsoft Office\Office15\WINWORD.EXE",
    r"C:\Program Files (x86)\Microsoft Office\Office15\WINWORD.EXE",
    r"C:\Program Files\Microsoft Office\Office14\WINWORD.EXE",
]

_SOFFICE_CANDIDATES = [
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    "/usr/bin/libreoffice",
    "/usr/bin/soffice",
    "/usr/local/bin/soffice",
    "/snap/bin/libreoffice",
]


def _from_registry() -> Optional[str]:
    """从注册表读 Word 的安装路径（比硬编码路径更可靠）。"""
    if not IS_WIN:
        return None
    try:
        import winreg
    except ImportError:
        return None
    for root in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        for key in (
            r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE",
            r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE",
        ):
            try:
                with winreg.OpenKey(root, key) as k:
                    path, _ = winreg.QueryValueEx(k, "")
                    if path and Path(path).exists():
                        return str(Path(path))
            except OSError:
                continue
    return None


def find_word() -> Optional[str]:
    """返回 WINWORD.EXE 的绝对路径，找不到返回 None。"""
    if not IS_WIN:
        return None
    got = shutil.which("WINWORD.EXE") or shutil.which("winword")
    if got:
        return got
    got = _from_registry()
    if got:
        return got
    for p in _WINWORD_CANDIDATES:
        if Path(p).exists():
            return p
    return None


def find_soffice() -> Optional[str]:
    """返回 soffice 可执行文件路径。"""
    for name in ("soffice", "soffice.exe", "libreoffice"):
        got = shutil.which(name)
        if got:
            return got
    for p in _SOFFICE_CANDIDATES:
        if Path(p).exists():
            return p
    return None


def has_module(name: str) -> bool:
    import importlib.util

    try:
        return importlib.util.find_spec(name) is not None
    except (ImportError, ValueError):
        return False


def pywin32_ready() -> tuple[bool, str]:
    """pywin32 装了但 DLL 没注册时会导入失败，这里一次性探完。"""
    if not IS_WIN:
        return False, "仅 Windows 支持 COM 自动化"
    try:
        import win32com.client  # noqa: F401
        import pythoncom  # noqa: F401
    except Exception as exc:  # pragma: no cover - 环境相关
        return False, f"pywin32 不可用：{exc}"
    return True, "pywin32 就绪"


def describe_environment() -> list[tuple[str, str, bool]]:
    """给 CLI / GUI 用的环境体检表：[(项目, 说明, 是否可用)]"""
    rows: list[tuple[str, str, bool]] = []
    word = find_word()
    rows.append(("Microsoft Word", word or "未安装", bool(word)))
    soffice = find_soffice()
    rows.append(("LibreOffice", soffice or "未安装", bool(soffice)))
    ready, msg = pywin32_ready()
    rows.append(("pywin32 (COM)", msg, ready))
    rows.append(("python-docx", "已安装" if has_module("docx") else "未安装", has_module("docx")))
    rows.append(("PyMuPDF (fitz)", "已安装" if has_module("fitz") else "未安装", has_module("fitz")))
    rows.append(("pdf2docx", "已安装" if has_module("pdf2docx") else "未安装", has_module("pdf2docx")))
    rows.append(("Python", f"{sys.version.split()[0]}  {sys.executable}", True))
    return rows
