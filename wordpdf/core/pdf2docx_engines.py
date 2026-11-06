"""PDF -> Word 的两个引擎。

1. Pdf2DocxEngine   —— pdf2docx：PyMuPDF 逐页解析文本块/表格/图片的版面坐标，再用 python-docx 重建
2. TextFlowEngine   —— PyMuPDF 抽取文本块，按阅读顺序写成一栏 docx（无版面，保内容）
"""

from __future__ import annotations

from pathlib import Path

from . import discovery
from .base import BaseEngine
from .model import ConversionError, Options


#: Symbol/Wingdings 私有区字符 → 常规符号（PDF 里的项目符号常被抽成这样）
_BULLET_MAP = {
    "\uf0b7": "•", "\uf0a7": "▪", "\uf0d8": "➢", "\uf06c": "●",
    "\uf0fc": "✓", "\uf0e0": "→", "\uf02d": "-",
}
_SYMBOL_FONTS = {"symbol", "symbolmt", "wingdings", "wingdings 2", "wingdings 3", "webdings"}


def _normalize_bullets(path: Path) -> None:
    """把 pdf2docx 产出的私有区符号换成人能看的字符，失败不影响主流程。"""
    try:
        from docx import Document
        from docx.oxml.ns import qn

        doc = Document(str(path))
        table = str.maketrans(_BULLET_MAP)
        changed = False

        def fix(paragraphs) -> None:
            nonlocal changed
            for para in paragraphs:
                for run in para.runs:
                    new = run.text.translate(table)
                    if new != run.text:
                        run.text = new
                        changed = True
                    # 符号字体（Symbol/Wingdings）里的字符在普通字体下会显示成方框，
                    # 解除字体绑定让它回落到正文默认字体
                    if (run.font.name or "").strip().lower() in _SYMBOL_FONTS:
                        rpr = run._element.rPr
                        if rpr is not None:
                            for el in rpr.findall(qn("w:rFonts")):
                                rpr.remove(el)
                                changed = True

        fix(doc.paragraphs)
        for tbl in doc.tables:
            for row in tbl.rows:
                for cell in row.cells:
                    fix(cell.paragraphs)
        if changed:
            doc.save(str(path))
    except Exception:
        pass


class Pdf2DocxEngine(BaseEngine):
    name = "pdf2docx"
    label = "pdf2docx 版面重建"
    note = "解析 PDF 的文本块/表格/图片坐标，映射成 Word 段落、表格与浮动图片"
    suffixes = {".pdf"}

    def probe(self) -> tuple[bool, str]:
        if not discovery.has_module("pdf2docx"):
            return False, "未安装 pdf2docx"
        if not discovery.has_module("fitz"):
            return False, "缺少 PyMuPDF"
        return True, "版面重建可用（表格/图片会被还原）"

    def convert(self, src: Path, dst: Path, opts: Options) -> None:
        try:
            from pdf2docx import Converter
        except Exception as exc:
            raise ConversionError(f"pdf2docx 导入失败：{exc}") from exc

        start, end = opts.page_range()
        cv = None
        try:
            cv = Converter(str(src))
            cv.convert(str(dst), start=start, end=end, multi_processing=bool(opts.multi_processing))
        except Exception as exc:
            raise ConversionError(f"版面重建失败：{exc}") from exc
        finally:
            if cv is not None:
                try:
                    cv.close()
                except Exception:
                    pass

        if not dst.exists() or dst.stat().st_size == 0:
            raise ConversionError("pdf2docx 未产出 docx")
        _normalize_bullets(dst)


class TextFlowEngine(BaseEngine):
    name = "text-flow"
    label = "PyMuPDF 纯文本流重建（兜底）"
    note = "只按阅读顺序搬文字和图片，不还原表格线与分栏；版面会重排"
    suffixes = {".pdf"}

    def probe(self) -> tuple[bool, str]:
        if not discovery.has_module("fitz"):
            return False, "未安装 PyMuPDF"
        if not discovery.has_module("docx"):
            return False, "未安装 python-docx"
        return True, "纯文本流可用"

    def convert(self, src: Path, dst: Path, opts: Options) -> None:
        try:
            import pymupdf as fitz  # type: ignore
        except ImportError:
            import fitz  # type: ignore
        from docx import Document

        start, end = opts.page_range()
        try:
            pdf = fitz.open(str(src))
        except Exception as exc:
            raise ConversionError(f"无法打开 PDF：{exc}") from exc

        doc = Document()
        try:
            last = pdf.page_count if end is None else min(end, pdf.page_count)
            for pno in range(start, last):
                page = pdf[pno]
                blocks = page.get_text("blocks") or []
                blocks.sort(key=lambda b: (round(b[1] / 6), b[0]))
                for block in blocks:
                    text = (block[4] or "").strip()
                    if not text:
                        continue
                    for line in text.splitlines():
                        line = line.strip()
                        if line:
                            doc.add_paragraph(line)
                if pno < last - 1:
                    doc.add_page_break()
            doc.save(str(dst))
        except Exception as exc:
            raise ConversionError(f"重建 docx 失败：{exc}") from exc
        finally:
            try:
                pdf.close()
            except Exception:
                pass
