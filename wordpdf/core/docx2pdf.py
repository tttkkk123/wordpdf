"""Word -> PDF 的三个引擎。

1. WordComEngine        —— 调用本机 Microsoft Word 的 ExportAsFixedFormat（排版 1:1 保真）
2. LibreOfficeEngine    —— 调用 soffice --headless --convert-to pdf（跨平台、免 Office）
3. PythonRendererEngine —— python-docx 解析 + PyMuPDF 重新排版（无任何 Office 依赖，降级可用）
"""

from __future__ import annotations

import os
import shutil
import subprocess
import tempfile
import time
from pathlib import Path

from . import discovery
from .base import BaseEngine
from .model import ConversionError, Options, WORD_SUFFIXES


# ---------------------------------------------------------------------------
# 引擎 1：Microsoft Word COM 自动化
# ---------------------------------------------------------------------------
class WordComEngine(BaseEngine):
    name = "word-com"
    label = "Microsoft Word COM 自动化"
    note = "直接驱动本机 Word 的 PDF 导出器，字体/分页/页眉页脚与原文完全一致"
    suffixes = WORD_SUFFIXES

    def probe(self) -> tuple[bool, str]:
        if os.name != "nt":
            return False, "仅 Windows 可用"
        ready, msg = discovery.pywin32_ready()
        if not ready:
            return False, msg
        exe = discovery.find_word()
        if not exe:
            return False, "未检测到 Microsoft Word"
        return True, f"Word 已就绪（{exe}）"

    def convert(self, src: Path, dst: Path, opts: Options) -> None:
        import pythoncom
        import win32com.client as win32

        pythoncom.CoInitialize()
        app = None
        owned = False
        started = time.time()
        try:
            # DispatchEx 起独立进程：不干扰用户正开着的文档，退出时可以安全 Quit
            try:
                app = win32.DispatchEx("Word.Application")
                owned = True
            except Exception:
                app = win32.GetActiveObject("Word.Application")
                owned = False

            try:
                app.Visible = False
            except Exception:
                pass
            try:
                app.DisplayAlerts = 0  # wdAlertsNone
            except Exception:
                pass

            doc = app.Documents.Open(
                FileName=str(src),
                ConfirmConversions=False,
                ReadOnly=True,
                AddToRecentFiles=False,
                Visible=False,
            )
            try:
                doc.ExportAsFixedFormat(
                    OutputFileName=str(dst),
                    ExportFormat=17,                            # wdExportFormatPDF
                    OpenAfterExport=False,
                    OptimizeFor=opts.optimize_for,              # 0 打印 / 1 屏幕
                    Range=0,                                    # wdExportAllDocument
                    Item=0,                                     # wdExportDocumentContent
                    IncludeDocProps=True,
                    KeepIRM=True,
                    CreateBookmarks=1 if opts.bookmarks else 0,  # 标题转书签
                    DocStructureTags=True,
                    BitmapMissingFonts=True,
                    UseISO19005_1=bool(opts.pdfa),
                )
            finally:
                try:
                    doc.Close(SaveChanges=0)
                except Exception:
                    pass
        except Exception as exc:
            raise ConversionError(f"Word 自动化失败：{exc}") from exc
        finally:
            if owned and app is not None:
                try:
                    app.Quit()
                except Exception:
                    pass
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

        if not dst.exists() or dst.stat().st_size == 0:
            raise ConversionError("Word 未生成 PDF 文件（耗时 %.1fs）" % (time.time() - started))


# ---------------------------------------------------------------------------
# 引擎 2：LibreOffice 无头转换
# ---------------------------------------------------------------------------
class LibreOfficeEngine(BaseEngine):
    name = "libreoffice"
    label = "LibreOffice 无头转换"
    note = "调用 soffice --headless，免装 Office，排版基本一致（复杂域/宏可能走样）"
    suffixes = WORD_SUFFIXES

    def probe(self) -> tuple[bool, str]:
        exe = discovery.find_soffice()
        if not exe:
            return False, "未检测到 LibreOffice"
        return True, f"soffice 已就绪（{exe}）"

    def convert(self, src: Path, dst: Path, opts: Options) -> None:
        exe = discovery.find_soffice()
        if not exe:
            raise ConversionError("未检测到 LibreOffice")
        with tempfile.TemporaryDirectory(prefix="wordpdf_lo_") as tmp:
            cmd = [
                exe,
                "--headless",
                "--norestore",
                "--nolockcheck",
                "--nodefault",
                f"-env:UserInstallation={Path(tmp) / 'profile'}",
                "--convert-to",
                "pdf:writer_pdf_Export",
                "--outdir",
                tmp,
                str(src),
            ]
            try:
                done = subprocess.run(
                    cmd,
                    capture_output=True,
                    timeout=180,
                    creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
                )
            except subprocess.TimeoutExpired as exc:
                raise ConversionError("LibreOffice 转换超时（180s）") from exc
            except OSError as exc:
                raise ConversionError(f"无法启动 LibreOffice：{exc}") from exc

            produced = Path(tmp) / (src.stem + ".pdf")
            if not produced.exists():
                err = (done.stderr or b"").decode("utf-8", "ignore").strip()
                raise ConversionError(f"LibreOffice 未产出 PDF：{err[:200] or done.returncode}")
            shutil.move(str(produced), str(dst))


# ---------------------------------------------------------------------------
# 引擎 3：纯 Python 兜底（不依赖任何 Office 软件）
# ---------------------------------------------------------------------------
class PythonRendererEngine(BaseEngine):
    name = "python-render"
    label = "纯 Python 重新排版（兜底）"
    note = "python-docx 解析正文结构 + PyMuPDF 重绘页面；只保证内容正确，复杂版式会简化"
    suffixes = WORD_SUFFIXES

    def probe(self) -> tuple[bool, str]:
        if not discovery.has_module("docx"):
            return False, "缺少 python-docx"
        if not discovery.has_module("fitz"):
            return False, "缺少 PyMuPDF"
        return True, "无需 Office，可直接排版"

    def convert(self, src: Path, dst: Path, opts: Options) -> None:
        if not self.probe()[0]:
            raise ConversionError(self.probe()[1])
        if src.suffix.lower() in {".doc", ".rtf", ".wps", ".odt"}:
            raise ConversionError("纯 Python 引擎只认 .docx，请先另存为 .docx")
        from .render_docx import render

        render(src, dst, opts)
