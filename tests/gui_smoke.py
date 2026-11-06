"""GUI 冒烟测试：自动点一次「开始转换」，把界面截下来并检查日志。

用法：
    .venv\\Scripts\\python.exe tests\\gui_smoke.py
产物：tests/_out/gui_smoke.png
"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import win32con      # noqa: E402
import win32gui      # noqa: E402
import win32ui       # noqa: E402
from PIL import Image  # noqa: E402

from wordpdf import gui as gui_mod  # noqa: E402

OUT = ROOT / "tests" / "_out"


def capture(hwnd: int, path: Path) -> tuple[int, int]:
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    w, h = right - left, bottom - top
    hwnd_dc = win32gui.GetWindowDC(hwnd)
    mfc_dc = win32ui.CreateDCFromHandle(hwnd_dc)
    save_dc = mfc_dc.CreateCompatibleDC()
    bmp = win32ui.CreateBitmap()
    bmp.CreateCompatibleBitmap(mfc_dc, w, h)
    save_dc.SelectObject(bmp)
    save_dc.BitBlt((0, 0), (w, h), mfc_dc, (0, 0), win32con.SRCCOPY)
    info = bmp.GetInfo()
    bits = bmp.GetBitmapBits(True)
    Image.frombuffer("RGB", (info["bmWidth"], info["bmHeight"]), bits, "raw", "BGRX", 0, 1).save(str(path))
    win32gui.DeleteObject(bmp.GetHandle())
    save_dc.DeleteDC()
    mfc_dc.DeleteDC()
    win32gui.ReleaseDC(hwnd, hwnd_dc)
    return w, h


def main() -> int:
    OUT.mkdir(parents=True, exist_ok=True)
    # 不弹模态框，方便无人值守
    gui_mod.messagebox.showinfo = lambda *a, **k: print("  [对话框]", a[0] if a else "")
    gui_mod.messagebox.showwarning = lambda *a, **k: print("  [警告框]", a[0] if a else "")
    gui_mod.messagebox.showerror = lambda *a, **k: print("  [错误框]", a[0] if a else "")

    gui_mod._enable_dpi_awareness()
    app = gui_mod.App()
    app.update()

    # 干净 clone 下先确保样例存在（不依赖是否跑过 selftest）
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    from selftest import make_pdf_sample, make_sample

    docx_sample = OUT / "样例报告.docx"
    pdf_sample = OUT / "样例合同.pdf"
    if not docx_sample.exists():
        make_sample(docx_sample)
    if not pdf_sample.exists():
        make_pdf_sample(pdf_sample)

    app._add_paths([docx_sample.resolve(), pdf_sample.resolve()])

    state = {"captured": False, "ticks": 0}

    def tick() -> None:
        state["ticks"] += 1
        text = app.log.get("1.0", "end").strip()
        if not state["captured"] and (("全部结束" in text) or state["ticks"] > 120):
            hwnd = win32gui.GetAncestor(app.winfo_id(), 2)
            print("  截图尺寸:", capture(hwnd, OUT / "gui_smoke.png"))
            print("  状态栏:", app.var_status.get())
            print("  日志尾部:", text.splitlines()[-1] if text else "(空)")
            state["captured"] = True
            app.after(400, app.destroy)
            return
        app.after(500, tick)

    app.after(800, app.start)
    app.after(1500, tick)
    app.mainloop()

    ok = state["captured"]
    print("  结果:", "OK" if ok else "超时")
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
