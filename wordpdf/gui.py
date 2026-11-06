"""图形界面（Tkinter）：拖入文件 → 选方向 → 转换。

界面只负责收集参数和展示日志，真正的活儿都交给 core.converter。
"""

from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from . import __version__
from .core import convert_many
from .core.discovery import describe_environment
from .core.model import WORD_SUFFIXES, PDF_SUFFIXES, Options

# --- 配色（浅色主题） ------------------------------------------------------
BG = "#f5f6f8"
CARD = "#ffffff"
BORDER = "#e2e5ea"
TEXT = "#1f2328"
MUTED = "#6b7280"
PRIMARY = "#2f6feb"
PRIMARY_DARK = "#2559c9"
OK_GREEN = "#16a34a"
BAD_RED = "#dc2626"

FONT = "Microsoft YaHei UI"
SUPPORTED = tuple(sorted(WORD_SUFFIXES | PDF_SUFFIXES))


def _enable_dpi_awareness() -> None:
    """必须在创建窗口之前调用，否则 150% 缩放下窗口会被裁切。"""
    if os.name != "nt":
        return
    try:
        import ctypes

        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(2)  # 每显示器 DPI 感知
        except Exception:
            ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        # 按系统 DPI 放大窗口尺寸，同时让 Tk 按 DPI 计算字号
        try:
            dpi = float(self.winfo_fpixels("1i"))
        except Exception:
            dpi = 96.0
        self.scale = max(1.0, dpi / 96.0)
        if self.scale > 1.01:
            try:
                self.tk.call("tk", "scaling", dpi / 72.0)
            except Exception:
                pass

        self.title(f"Word ⇄ PDF 转换器  v{__version__}")
        self.geometry(f"{int(1000 * self.scale)}x{int(845 * self.scale)}")
        self.minsize(int(900 * self.scale), int(700 * self.scale))
        self.configure(bg=BG)

        self.files: list[Path] = []
        self.q: "queue.Queue[tuple]" = queue.Queue()
        self.busy = False
        self.last_outdir: Path | None = None

        self.var_dir = tk.StringVar(value="auto")
        self.var_out = tk.StringVar(value="")
        self.var_overwrite = tk.BooleanVar(value=False)
        self.var_pdfa = tk.BooleanVar(value=False)
        self.var_screen = tk.BooleanVar(value=False)
        self.var_pages = tk.StringVar(value="")
        self.var_status = tk.StringVar(value="就绪")

        self._init_style()
        self._build()
        self._init_dnd()
        self._refresh_engine_badge()
        self.after(120, self._pump)

    # ------------------------------------------------------------------ 样式
    def _init_style(self) -> None:
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure(".", font=(FONT, 10), background=BG, foreground=TEXT)
        style.configure("Card.TFrame", background=CARD)
        style.configure("Bg.TFrame", background=BG)
        style.configure("TLabel", background=BG, foreground=TEXT)
        style.configure("Card.TLabel", background=CARD, foreground=TEXT)
        style.configure("Muted.TLabel", background=BG, foreground=MUTED, font=(FONT, 9))
        style.configure("CardMuted.TLabel", background=CARD, foreground=MUTED, font=(FONT, 9))
        style.configure("Title.TLabel", background=BG, foreground=TEXT, font=(FONT, 17, "bold"))
        style.configure("H2.TLabel", background=CARD, foreground=TEXT, font=(FONT, 10, "bold"))
        style.configure("TCheckbutton", background=CARD, foreground=TEXT)
        style.configure("TRadiobutton", background=CARD, foreground=TEXT)
        style.configure("Accent.TButton", font=(FONT, 11, "bold"), padding=(18, 9))
        style.map(
            "Accent.TButton",
            background=[("!disabled", PRIMARY), ("pressed", PRIMARY_DARK), ("disabled", "#b9c6e4")],
            foreground=[("!disabled", "white"), ("disabled", "#f0f3fa")],
        )
        style.configure("Flat.TButton", padding=(10, 5))
        style.configure("TEntry", padding=5)
        style.configure("Horizontal.TProgressbar", background=PRIMARY, troughcolor="#e8ecf3", thickness=8)

    # ------------------------------------------------------------------ 布局
    def _build(self) -> None:
        root = ttk.Frame(self, style="Bg.TFrame", padding=(18, 14, 18, 12))
        root.pack(fill="both", expand=True)
        root.columnconfigure(0, weight=3, minsize=520)
        root.columnconfigure(1, weight=2, minsize=300)
        root.rowconfigure(1, weight=3)
        root.rowconfigure(3, weight=2)

        # 头部
        head = ttk.Frame(root, style="Bg.TFrame")
        head.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 12))
        head.columnconfigure(1, weight=1)
        ttk.Label(head, text="Word ⇄ PDF 转换器", style="Title.TLabel").grid(row=0, column=0, sticky="w")
        self.badge = tk.Label(head, text="", bg=BG, fg=MUTED, font=(FONT, 9))
        self.badge.grid(row=2, column=0, columnspan=2, sticky="w", pady=(3, 0))
        ttk.Label(
            head,
            text="混合引擎：Word COM 保真导出 · pdf2docx 版面重建 · 失败自动降级",
            style="Muted.TLabel",
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(2, 0))

        # 左：文件区
        left = tk.Frame(root, bg=CARD, highlightbackground=BORDER, highlightthickness=1)
        left.grid(row=1, column=0, sticky="nsew", padx=(0, 12))
        left.rowconfigure(1, weight=1)
        left.columnconfigure(0, weight=1)

        bar = tk.Frame(left, bg=CARD)
        bar.grid(row=0, column=0, sticky="ew", padx=12, pady=(10, 6))
        bar.columnconfigure(1, weight=1)
        ttk.Label(bar, text="待转换文件", style="H2.TLabel").grid(row=0, column=0, sticky="w")
        self.count_lbl = ttk.Label(bar, text="0 个", style="CardMuted.TLabel")
        self.count_lbl.grid(row=0, column=1, sticky="e", padx=(0, 8))
        ttk.Button(bar, text="添加", style="Flat.TButton", command=self.add_files).grid(row=0, column=2)
        ttk.Button(bar, text="清空", style="Flat.TButton", command=self.clear_files).grid(row=0, column=3, padx=(6, 0))

        wrap = tk.Frame(left, bg=CARD)
        wrap.grid(row=1, column=0, sticky="nsew", padx=12)
        wrap.rowconfigure(0, weight=1)
        wrap.columnconfigure(0, weight=1)
        self.listbox = tk.Listbox(
            wrap, bg="#fbfcfe", fg=TEXT, font=(FONT, 10), relief="flat",
            highlightthickness=1, highlightbackground=BORDER, selectbackground="#dbe7ff",
            selectforeground=TEXT, activestyle="none", bd=0,
        )
        self.listbox.grid(row=0, column=0, sticky="nsew")
        sb = ttk.Scrollbar(wrap, orient="vertical", command=self.listbox.yview)
        sb.grid(row=0, column=1, sticky="ns")
        self.listbox.configure(yscrollcommand=sb.set)
        self.listbox.bind("<Delete>", lambda _e: self.remove_selected())
        self.empty_hint = tk.Label(
            wrap, text="把 .docx / .pdf 文件拖到这里\n或点右上角「添加」",
            bg="#fbfcfe", fg="#9aa3af", font=(FONT, 10), justify="center",
        )
        self.empty_hint.place(relx=0.5, rely=0.42, anchor="center")

        foot = tk.Frame(left, bg=CARD)
        foot.grid(row=2, column=0, sticky="ew", padx=12, pady=(6, 10))
        ttk.Button(foot, text="移除选中", style="Flat.TButton", command=self.remove_selected).pack(side="left")
        self.out_hint = ttk.Label(foot, text="输出：源文件同目录", style="CardMuted.TLabel")
        self.out_hint.pack(side="right")

        # 右：参数区
        right = tk.Frame(root, bg=CARD, highlightbackground=BORDER, highlightthickness=1)
        right.grid(row=1, column=1, sticky="nsew")
        right.columnconfigure(0, weight=1)
        ttk.Label(right, text="转换设置", style="H2.TLabel").grid(row=0, column=0, sticky="w", padx=12, pady=(10, 8))

        box = tk.Frame(right, bg=CARD)
        box.grid(row=1, column=0, sticky="nsew", padx=12)
        box.columnconfigure(0, weight=1)

        ttk.Label(box, text="转换方向", style="CardMuted.TLabel").grid(row=0, column=0, sticky="w")
        dirs = tk.Frame(box, bg=CARD)
        dirs.grid(row=1, column=0, sticky="ew", pady=(2, 10))
        for i, (val, label) in enumerate((("auto", "自动识别"), ("docx2pdf", "Word→PDF"), ("pdf2docx", "PDF→Word"))):
            ttk.Radiobutton(dirs, text=label, value=val, variable=self.var_dir,
                            command=self._refresh_engine_badge).grid(row=0, column=i, sticky="w", padx=(0, 14))

        ttk.Label(box, text="输出目录", style="CardMuted.TLabel").grid(row=2, column=0, sticky="w")
        orow = tk.Frame(box, bg=CARD)
        orow.grid(row=3, column=0, sticky="ew", pady=(2, 10))
        orow.columnconfigure(0, weight=1)
        ttk.Entry(orow, textvariable=self.var_out).grid(row=0, column=0, sticky="ew")
        ttk.Button(orow, text="…", width=3, command=self.pick_outdir).grid(row=0, column=1, padx=(4, 0))

        ttk.Label(box, text="选项", style="CardMuted.TLabel").grid(row=4, column=0, sticky="w")
        ttk.Checkbutton(box, text="覆盖同名输出文件", variable=self.var_overwrite).grid(row=5, column=0, sticky="w")
        ttk.Checkbutton(box, text="PDF/A-1b 归档格式（仅 Word→PDF）", variable=self.var_pdfa).grid(row=6, column=0, sticky="w")
        ttk.Checkbutton(box, text="按屏幕质量导出，体积更小", variable=self.var_screen).grid(row=7, column=0, sticky="w")

        prow = tk.Frame(box, bg=CARD)
        prow.grid(row=8, column=0, sticky="ew", pady=(8, 0))
        ttk.Label(prow, text="页码范围", style="CardMuted.TLabel").pack(side="left")
        ttk.Entry(prow, textvariable=self.var_pages, width=10).pack(side="left", padx=(6, 6))
        ttk.Label(prow, text="如 1-8", style="CardMuted.TLabel").pack(side="left")

        ttk.Separator(box, orient="horizontal").grid(row=9, column=0, sticky="ew", pady=12)
        ttk.Label(box, text="引擎状态", style="CardMuted.TLabel").grid(row=10, column=0, sticky="w")
        self.engine_box = tk.Frame(box, bg=CARD)
        self.engine_box.grid(row=11, column=0, sticky="ew", pady=(4, 0))
        self.engine_box.columnconfigure(2, weight=1)
        ttk.Button(box, text="环境体检", style="Flat.TButton", command=self.show_env).grid(row=12, column=0, sticky="w", pady=(10, 0))

        # 中部：进度 + 开始
        act = ttk.Frame(root, style="Bg.TFrame")
        act.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(12, 8))
        act.columnconfigure(0, weight=1)
        self.pb = ttk.Progressbar(act, mode="determinate", maximum=100)
        self.pb.grid(row=0, column=0, sticky="ew", padx=(0, 12))
        self.btn_run = ttk.Button(act, text="开始转换", style="Accent.TButton", command=self.start)
        self.btn_run.grid(row=0, column=1)
        ttk.Button(act, text="打开输出目录", style="Flat.TButton", command=self.open_outdir).grid(row=0, column=2, padx=(8, 0))
        tk.Label(act, textvariable=self.var_status, bg=BG, fg=MUTED, font=(FONT, 9)).grid(
            row=1, column=0, sticky="w", pady=(4, 0)
        )

        # 日志
        logbox = tk.Frame(root, bg=CARD, highlightbackground=BORDER, highlightthickness=1)
        logbox.grid(row=3, column=0, columnspan=2, sticky="nsew")
        logbox.rowconfigure(1, weight=1)
        logbox.columnconfigure(0, weight=1)
        ttk.Label(logbox, text="运行日志", style="H2.TLabel").grid(row=0, column=0, sticky="w", padx=12, pady=(8, 4))
        self.log = tk.Text(
            logbox, height=6, bg="#fbfcfe", fg="#374151", font=("Consolas", 9),
            relief="flat", bd=0, highlightthickness=0, wrap="word", padx=10, pady=6,
        )
        self.log.grid(row=1, column=0, sticky="nsew", padx=(8, 0), pady=(0, 8))
        lsb = ttk.Scrollbar(logbox, orient="vertical", command=self.log.yview)
        lsb.grid(row=1, column=1, sticky="ns", pady=(0, 8), padx=(0, 4))
        self.log.configure(yscrollcommand=lsb.set, state="disabled")
        self.log.tag_config("ok", foreground=OK_GREEN)
        self.log.tag_config("err", foreground=BAD_RED)
        self.log.tag_config("dim", foreground=MUTED)

    # ------------------------------------------------------------ 拖拽支持
    def _init_dnd(self) -> None:
        try:
            from tkinterdnd2 import DND_FILES  # type: ignore

            self.listbox.drop_target_register(DND_FILES)
            self.listbox.dnd_bind("<<Drop>>", self._on_drop)
            self._dnd = True
        except Exception:
            self._dnd = False

    def _on_drop(self, event) -> None:
        paths = self.tk.splitlist(event.data)
        self._add_paths([Path(p) for p in paths])

    # ------------------------------------------------------------ 文件操作
    def add_files(self) -> None:
        paths = filedialog.askopenfilenames(
            title="选择要转换的文件",
            filetypes=[
                ("Word 与 PDF", "*.docx *.doc *.docm *.rtf *.pdf"),
                ("Word 文档", "*.docx *.doc *.docm *.rtf"),
                ("PDF 文件", "*.pdf"),
                ("全部文件", "*.*"),
            ],
        )
        self._add_paths([Path(p) for p in paths])

    def _add_paths(self, paths: list[Path]) -> None:
        added = 0
        for p in paths:
            if not p.is_file():
                continue
            if p.suffix.lower() not in SUPPORTED:
                self._say(f"跳过不支持的类型：{p.name}", "dim")
                continue
            if p.resolve() in [f.resolve() for f in self.files]:
                continue
            self.files.append(p)
            self.listbox.insert("end", f"  {p.name}    ·  {p.parent.name}")
            added += 1
        if added:
            self.empty_hint.place_forget()
            self.count_lbl.configure(text=f"{len(self.files)} 个")
            self.var_status.set(f"已加入 {added} 个文件")
            self._refresh_engine_badge()

    def remove_selected(self) -> None:
        for idx in reversed(self.listbox.curselection()):
            self.listbox.delete(idx)
            del self.files[idx]
        self.count_lbl.configure(text=f"{len(self.files)} 个")
        if not self.files:
            self.empty_hint.place(relx=0.5, rely=0.42, anchor="center")

    def clear_files(self) -> None:
        self.files.clear()
        self.listbox.delete(0, "end")
        self.count_lbl.configure(text="0 个")
        self.empty_hint.place(relx=0.5, rely=0.42, anchor="center")

    def pick_outdir(self) -> None:
        d = filedialog.askdirectory(title="选择输出目录")
        if d:
            self.var_out.set(d)
            self.out_hint.configure(text=f"输出：{d}")

    def open_outdir(self) -> None:
        target = self.last_outdir or (Path(self.var_out.get()) if self.var_out.get() else None)
        if not target or not Path(target).exists():
            messagebox.showinfo("提示", "还没有输出目录，先转换一次吧")
            return
        try:
            if os.name == "nt":
                os.startfile(str(target))  # type: ignore[attr-defined]
            else:
                subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", str(target)])
        except Exception as exc:
            messagebox.showerror("打不开目录", str(exc))

    # ------------------------------------------------------------ 引擎状态
    def _refresh_engine_badge(self) -> None:
        from .core import engine_table

        for child in self.engine_box.winfo_children():
            child.destroy()
        dirs = ["docx2pdf", "pdf2docx"] if self.var_dir.get() == "auto" else [self.var_dir.get()]
        titles = {"docx2pdf": "Word → PDF", "pdf2docx": "PDF → Word"}
        badges = []
        row = 0
        for d in dirs:
            engines = engine_table(d)
            primary = next((e for e in engines if e[2]), None)
            backups = len(engines) - 1 if primary else 0

            tk.Label(self.engine_box, text=titles[d], bg=CARD, fg=MUTED, font=(FONT, 9)).grid(
                row=row, column=0, sticky="w", pady=(2, 0)
            )
            if primary:
                tk.Label(self.engine_box, text="●", bg=CARD, fg=OK_GREEN, font=(FONT, 8)).grid(
                    row=row, column=1, sticky="w", padx=(8, 3), pady=(2, 0)
                )
                detail = primary[1] + (f"（另有 {backups} 级备用）" if backups else "")
                tk.Label(self.engine_box, text=detail, bg=CARD, fg=TEXT, font=(FONT, 9),
                         anchor="w", justify="left", wraplength=int(190 * self.scale)).grid(
                    row=row, column=2, sticky="w", pady=(2, 0)
                )
                badges.append(f"{titles[d]}: {primary[1]}")
            else:
                tk.Label(self.engine_box, text="●", bg=CARD, fg=BAD_RED, font=(FONT, 8)).grid(
                    row=row, column=1, sticky="w", padx=(8, 3), pady=(2, 0)
                )
                tk.Label(self.engine_box, text="无可用引擎", bg=CARD, fg=BAD_RED, font=(FONT, 9)).grid(
                    row=row, column=2, sticky="w", pady=(2, 0)
                )
                badges.append(f"{titles[d]}: 无可用引擎")
            row += 1

        short = [b.replace("Microsoft ", "").split("（")[0] for b in badges]
        text = "    ".join(short) + ("    ·    可拖拽文件" if self._dnd else "")
        self.badge.configure(text=text, fg=MUTED)

    def show_env(self) -> None:
        win = tk.Toplevel(self)
        win.title("环境体检")
        win.configure(bg=BG)
        win.geometry("620x420")
        ttk.Label(win, text="本机能力探测结果", style="Title.TLabel").pack(anchor="w", padx=16, pady=(14, 8))
        box = tk.Text(win, bg=CARD, fg=TEXT, font=("Consolas", 9), relief="flat", padx=12, pady=10, wrap="word")
        box.pack(fill="both", expand=True, padx=16, pady=(0, 16))
        lines = []
        for name, msg, ok in describe_environment():
            lines.append(f"[{'√' if ok else '×'}] {name:<18} {msg}")
        lines.append("")
        lines.append("引擎降级链")
        from .core import engine_table

        for d, title in (("docx2pdf", "Word → PDF"), ("pdf2docx", "PDF → Word")):
            lines.append(f"  {title}:")
            for i, (name, label, ok, msg) in enumerate(engine_table(d), 1):
                lines.append(f"    {i}. [{'可用' if ok else '不可用'}] {label} — {msg}")
        box.insert("1.0", "\n".join(lines))
        box.configure(state="disabled")

    # ------------------------------------------------------------ 执行转换
    def start(self) -> None:
        if self.busy:
            return
        if not self.files:
            messagebox.showinfo("还没有文件", "先添加几个 Word 或 PDF 文件")
            return
        outdir = self.var_out.get().strip() or None
        if outdir:
            try:
                Path(outdir).mkdir(parents=True, exist_ok=True)
            except Exception as exc:
                messagebox.showerror("输出目录不可用", str(exc))
                return
        self.busy = True
        self.btn_run.configure(state="disabled", text="转换中…")
        self.pb.configure(value=0, maximum=len(self.files))
        self.var_status.set("正在转换…")
        self._say(f"开始转换 {len(self.files)} 个文件", "dim")

        opts = Options(
            overwrite=self.var_overwrite.get(),
            optimize_for=1 if self.var_screen.get() else 0,
            pdfa=self.var_pdfa.get(),
            bookmarks=True,
            pages=self.var_pages.get().strip(),
        )
        files = list(self.files)
        direction = self.var_dir.get()

        threading.Thread(
            target=self._worker,
            args=(files, outdir, direction, opts),
            daemon=True,
        ).start()

    def _worker(self, files, outdir, direction, opts) -> None:
        def log(msg: str) -> None:
            self.q.put(("log", msg))

        def on_file(i, total, res) -> None:
            self.q.put(("progress", i, total, res))

        try:
            results = convert_many(
                files, outdir=outdir, direction=direction,
                options=opts, log=log, on_file=on_file,
            )
            self.q.put(("done", results, outdir))
        except Exception as exc:  # 兜底，界面不能因此崩
            self.q.put(("fatal", str(exc)))

    def _pump(self) -> None:
        try:
            while True:
                msg = self.q.get_nowait()
                kind = msg[0]
                if kind == "log":
                    text = msg[1]
                    tag = ""
                    if text.startswith("  使用引擎"):
                        tag = "dim"
                    self._say(text, tag)
                elif kind == "progress":
                    _, i, total, res = msg
                    self.pb.configure(value=i, maximum=total)
                    self.var_status.set(f"已处理 {i}/{total}")
                    self._say(res.summary, "ok" if res.ok else "err")
                elif kind == "done":
                    results, outdir = msg[1], msg[2]
                    self._finish(results, outdir)
                elif kind == "fatal":
                    self.busy = False
                    self.btn_run.configure(state="normal", text="开始转换")
                    self.var_status.set("出错了")
                    self._say(f"异常：{msg[1]}", "err")
                    messagebox.showerror("转换失败", str(msg[1]))
        except queue.Empty:
            pass
        self.after(120, self._pump)

    def _finish(self, results, outdir) -> None:
        self.busy = False
        self.btn_run.configure(state="normal", text="开始转换")
        ok = sum(1 for r in results if r.ok)
        if outdir:
            self.last_outdir = Path(outdir)
        elif results:
            self.last_outdir = results[0].dst.parent if results[0].dst else None
        self.var_status.set(f"完成：成功 {ok} / 共 {len(results)}")
        self.pb.configure(value=len(results))
        self._say(f"全部结束：成功 {ok}，失败 {len(results) - ok}", "ok" if ok == len(results) else "err")
        if ok == len(results):
            messagebox.showinfo("完成", f"{ok} 个文件转换完成\n输出目录：{self.last_outdir or '源文件目录'}")
        else:
            bad = [r for r in results if not r.ok]
            detail = "\n".join(f"· {r.src.name}：{r.message[:120]}" for r in bad[:6])
            messagebox.showwarning("部分失败", f"成功 {ok} 个，失败 {len(results) - ok} 个\n\n{detail}")

    # ------------------------------------------------------------ 日志
    def _say(self, text: str, tag: str = "") -> None:
        self.log.configure(state="normal")
        self.log.insert("end", text + "\n", tag or ())
        self.log.see("end")
        self.log.configure(state="disabled")


def main() -> None:
    _enable_dpi_awareness()
    # pythonw 启动时没有控制台，第三方库（如 pdf2docx）打日志会炸，这里兜一下
    if sys.stdout is None:
        sys.stdout = open(os.devnull, "w", encoding="utf-8")
    if sys.stderr is None:
        sys.stderr = open(os.devnull, "w", encoding="utf-8")
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
