"""GUI 启动脚本（双击 / pythonw 运行，出错不会静默消失）。

被 启动转换器.bat 调用；任何异常都会写日志并弹窗提示，绝不"一闪而过"。
"""

from __future__ import annotations

import os
import sys
import traceback
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))


def _write_log(text: str) -> Path | None:
    try:
        logdir = ROOT / "logs"
        logdir.mkdir(exist_ok=True)
        path = logdir / "startup.log"
        path.write_text(text, encoding="utf-8")
        return path
    except Exception:
        return None


def _show_error(detail: str, log_path: Path | None) -> None:
    """用最直接的方式告诉用户出了什么事。"""
    try:
        import tkinter as tk
        from tkinter import messagebox

        root = tk.Tk()
        root.withdraw()
        tail = detail.strip().splitlines()[-8:]
        msg = "启动失败，错误信息：\n\n" + "\n".join(tail)
        if log_path:
            msg += f"\n\n完整日志：{log_path}"
        messagebox.showerror("Word ⇄ PDF 转换器", msg)
        root.destroy()
    except Exception:
        pass


def main() -> int:
    # pythonw 下没有控制台，第三方库打印日志会炸，先兜住
    if sys.stdout is None:
        try:
            (ROOT / "logs").mkdir(exist_ok=True)
            sys.stdout = open(ROOT / "logs" / "stdout.log", "a", encoding="utf-8")
        except Exception:
            sys.stdout = open(os.devnull, "w", encoding="utf-8")
    if sys.stderr is None:
        sys.stderr = sys.stdout

    try:
        from wordpdf.gui import main as gui_main

        gui_main()
        return 0
    except SystemExit:
        raise
    except Exception:
        detail = traceback.format_exc()
        log_path = _write_log(detail)
        _show_error(detail, log_path)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
