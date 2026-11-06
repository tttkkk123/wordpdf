"""纯 Python 兜底排版：python-docx 读结构，PyMuPDF 画页面。

只保证"内容正确、版面整齐"，不追求与 Word 像素级一致 ——
它存在的意义是：机器上没装 Office 也能出 PDF。
"""

from __future__ import annotations

import re
from pathlib import Path

try:  # PyMuPDF 新旧包名兼容
    import pymupdf as fitz  # type: ignore
except ImportError:  # pragma: no cover
    import fitz  # type: ignore

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

from . import fonts
from .model import ConversionError, Options

EMU_PER_PT = 12700.0
LEADING = 1.62

_HEADING_SIZE = {1: 19.0, 2: 16.5, 3: 14.5, 4: 13.0, 5: 12.0, 6: 11.0}
_TITLE_STYLES = {"title", "标题", "文档标题"}
#: 不允许出现在行首的字符（简化版中文换行禁则）
_NO_START = set("%）］｝〕〉》」』】、。，．；：！？…％·,.;:!?)]}")
_ALIGN_LEFT, _ALIGN_CENTER, _ALIGN_RIGHT, _ALIGN_JUSTIFY = 0, 1, 2, 3
_HEADING_RE = re.compile(r"^(heading|标题|h)(\s*)([1-6])$", re.I)


def _text_len(text: str, size: float, spec: "fonts.FontSpec | None" = None) -> float:
    """量字宽：优先用嵌入字体的真实度量。"""
    return (spec or fonts.regular()).width(text, size)


def _draw(page, point, text: str, size: float, spec: "fonts.FontSpec", color) -> None:
    """在页面上写一行字；有粗体字体文件就用它，没有就靠微移重绘模拟。"""
    kwargs = dict(fontsize=size, color=color)
    if spec.file:
        kwargs.update(fontname=spec.alias, fontfile=spec.file)
    else:
        kwargs.update(fontname=spec.alias)
    page.insert_text(point, text, **kwargs)


class _Writer:
    def __init__(self, doc, page_w: float, page_h: float, margins: tuple[float, float, float, float]):
        self.out = fitz.open()
        self.page_w = page_w
        self.page_h = page_h
        self.ml, self.mr, self.mt, self.mb = margins
        self.width = page_w - self.ml - self.mr
        self.page = self._new_page()

    # -- 页面 ----------------------------------------------------------
    def _new_page(self):
        page = self.out.new_page(width=self.page_w, height=self.page_h)
        self.y = self.mt
        return page

    def ensure(self, height: float) -> None:
        if self.y + height > self.page_h - self.mb:
            self.page = self._new_page()

    # -- 文本 ----------------------------------------------------------
    def wrap(self, text: str, size: float, width: float, spec=None) -> list[str]:
        text = text.replace("\u00a0", " ").replace("\t", "    ")
        if not text:
            return [""]
        lines: list[str] = []
        cur, cur_w, i = "", 0.0, 0
        while i < len(text):
            ch = text[i]
            if ch == "\n":
                lines.append(cur)
                cur, cur_w = "", 0.0
                i += 1
                continue
            # 西文单词整体不断行
            if ch.isascii() and (ch.isalnum() or ch in "'-"):
                j = i
                while j < len(text) and text[j].isascii() and (text[j].isalnum() or text[j] in "'-."):
                    j += 1
                word = text[i:j]
                w = _text_len(word, size, spec)
                if cur_w + w > width and cur:
                    lines.append(cur)
                    cur, cur_w = word, w
                else:
                    cur, cur_w = cur + word, cur_w + w
                i = j
                continue
            w = _text_len(ch, size, spec)
            if cur_w + w > width and cur:
                # 中文换行禁则：标点不能出现在行首，把它前面的字一起挪下去
                if ch in _NO_START and len(cur) > 1:
                    moved = cur[-1]
                    lines.append(cur[:-1])
                    cur = moved + ch
                    cur_w = _text_len(cur, size, spec)
                    i += 1
                    continue
                lines.append(cur)
                cur, cur_w = "", 0.0
                if ch == " ":
                    i += 1
                    continue
            cur, cur_w = cur + ch, cur_w + w
            i += 1
        lines.append(cur)
        return lines

    def paragraph(
        self,
        text: str,
        size: float = 10.5,
        bold: bool = False,
        align: int = _ALIGN_LEFT,
        color: tuple[float, float, float] = (0, 0, 0),
        space_before: float = 0.0,
        space_after: float = 4.0,
        indent: float = 0.0,
    ) -> None:
        width = self.width - indent
        spec = fonts.bold() if bold else fonts.regular()
        lines = self.wrap(text, size, width, spec)
        line_h = size * LEADING
        self.y += space_before
        for line in lines:
            self.ensure(line_h)
            x = self.ml + indent
            if align == _ALIGN_CENTER:
                x = self.ml + indent + (width - _text_len(line, size, spec)) / 2
            elif align == _ALIGN_RIGHT:
                x = self.ml + indent + width - _text_len(line, size, spec)
            baseline = self.y + size * 1.06
            if line:
                _draw(self.page, fitz.Point(x, baseline), line, size, spec, color)
                if bold and not spec.file:  # 内置 CJK 字体无粗体，微移重绘
                    _draw(self.page, fitz.Point(x + 0.42, baseline), line, size, spec, color)
            self.y += line_h
        self.y += space_after

    # -- 表格 ----------------------------------------------------------
    def table(self, table: Table, base_size: float = 9.5) -> None:
        rows = table.rows
        if not rows:
            return
        cols = max(len(r.cells) for r in rows)
        if cols == 0:
            return
        col_w = self.width / cols
        pad = 4.0
        line_h = base_size * 1.5

        def row_layout(row) -> list[list[str]]:
            cells = []
            for idx in range(cols):
                try:
                    text = row.cells[idx].text.strip()
                except IndexError:
                    text = ""
                cells.append(self.wrap(text, base_size, col_w - 2 * pad))
            return cells

        first_row_is_header = True
        header_cells: list[list[str]] | None = None
        for r_idx, row in enumerate(rows):
            cells = row_layout(row)
            height = max(len(c) for c in cells) * line_h + 2 * pad
            if self.y + min(height, 2 * line_h) > self.page_h - self.mb:
                self.page = self._new_page()
                if header_cells is not None:  # 跨页重复表头
                    h_h = max(len(c) for c in header_cells) * line_h + 2 * pad
                    self._draw_table_row(header_cells, col_w, pad, base_size, line_h, h_h, header=True)
            if r_idx == 0 and first_row_is_header:
                header_cells = cells
            self._draw_table_row(cells, col_w, pad, base_size, line_h, height, header=(r_idx == 0 and first_row_is_header))
        self.y += 8

    def _draw_table_row(self, cells, col_w, pad, base_size, line_h, height, header=False):
        top = self.y
        x0 = self.ml
        spec = fonts.bold() if header else fonts.regular()
        self.page.draw_rect(
            fitz.Rect(x0, top, x0 + col_w * len(cells), top + height),
            color=(0.55, 0.55, 0.55), fill=(0.94, 0.94, 0.94) if header else None, width=0.6,
        )
        for c_idx, lines in enumerate(cells):
            cx = x0 + c_idx * col_w
            if c_idx:
                self.page.draw_line(fitz.Point(cx, top), fitz.Point(cx, top + height),
                                    color=(0.55, 0.55, 0.55), width=0.6)
            for l_idx, line in enumerate(lines):
                if not line:
                    continue
                point = fitz.Point(cx + pad, top + pad + base_size * 1.05 + l_idx * line_h)
                _draw(self.page, point, line, base_size, spec, (0, 0, 0))
                if header and not spec.file:
                    _draw(self.page, fitz.Point(point.x + 0.4, point.y), line, base_size, spec, (0, 0, 0))
        self.y = top + height

    # -- 图片 ----------------------------------------------------------
    def image(self, blob: bytes, w_pt: float, h_pt: float) -> None:
        max_w = self.width
        if w_pt <= 0 or h_pt <= 0:
            return
        scale = min(1.0, max_w / w_pt)
        w_pt, h_pt = w_pt * scale, h_pt * scale
        self.ensure(h_pt)
        rect = fitz.Rect(self.ml, self.y, self.ml + w_pt, self.y + h_pt)
        try:
            self.page.insert_image(rect, stream=blob)
        except Exception:
            return
        self.y += h_pt + 6

    # -- 收尾 ----------------------------------------------------------
    def save(self, dst: Path) -> None:
        total = self.out.page_count
        spec = fonts.regular()
        for idx in range(total):
            p = self.out[idx]
            label = f"- {idx + 1} / {total} -"
            size = 8.5
            w = _text_len(label, size, spec)
            _draw(
                p, fitz.Point((self.page_w - w) / 2, self.page_h - self.mb / 2),
                label, size, spec, (0.45, 0.45, 0.45),
            )
        try:
            self.out.subset_fonts()   # 只嵌入用到的字形，否则整份字体会把 PDF 撑到几十 MB
        except Exception:
            pass
        self.out.save(str(dst), garbage=4, deflate=True, deflate_fonts=True)
        self.out.close()


# ---------------------------------------------------------------------------
def _heading_level(para: Paragraph) -> int:
    style = (para.style.name or "").strip()
    m = _HEADING_RE.match(style)
    if m:
        return int(m.group(3))
    # 中文样式名常见 "标题 1"，也兼容 outlineLvl
    ppr = para._p.pPr
    if ppr is not None:
        lvl = ppr.find(qn("w:outlineLvl"))
        if lvl is not None:
            try:
                return min(6, int(lvl.get(qn("w:val"))) + 1)
            except (TypeError, ValueError):
                pass
    return 0


def _is_list_item(para: Paragraph) -> bool:
    ppr = para._p.pPr
    if ppr is not None and ppr.find(qn("w:numPr")) is not None:
        return True
    return "list" in (para.style.name or "").lower()


def _alignment(para: Paragraph) -> int:
    value = para.alignment
    if value is None:
        return _ALIGN_LEFT
    try:
        return {0: _ALIGN_LEFT, 1: _ALIGN_CENTER, 2: _ALIGN_RIGHT, 3: _ALIGN_JUSTIFY,
                4: _ALIGN_JUSTIFY, 5: _ALIGN_LEFT}.get(int(value), _ALIGN_LEFT)
    except (TypeError, ValueError):
        return _ALIGN_LEFT


def _iter_blocks(doc):
    for child in doc.element.body.iterchildren():
        if child.tag == qn("w:p"):
            yield Paragraph(child, doc)
        elif child.tag == qn("w:tbl"):
            yield Table(child, doc)


def _para_images(para: Paragraph):
    """产出 (blob, 宽pt, 高pt)。"""
    for blip in para._p.findall(".//" + qn("a:blip")):
        rid = blip.get(qn("r:embed"))
        if not rid:
            continue
        try:
            part = para.part.related_parts[rid]
            blob = part.blob
        except Exception:
            continue
        # 尺寸取 drawing 的 wp:extent
        w_pt = h_pt = 0.0
        extent = para._p.find(".//" + qn("wp:extent"))
        if extent is not None:
            try:
                w_pt = int(extent.get("cx")) / EMU_PER_PT
                h_pt = int(extent.get("cy")) / EMU_PER_PT
            except (TypeError, ValueError):
                pass
        if w_pt <= 0:
            try:
                px = fitz.Pixmap(blob)
                w_pt, h_pt = px.width * 0.75, px.height * 0.75
            except Exception:
                w_pt, h_pt = 300.0, 200.0
        yield blob, w_pt, h_pt


def render(src: Path, dst: Path, opts: Options) -> None:
    try:
        doc = Document(str(src))
    except Exception as exc:
        raise ConversionError(f"无法读取 docx：{exc}") from exc

    section = doc.sections[0] if doc.sections else None
    if section is not None:
        page_w = (section.page_width or 7560310) / EMU_PER_PT
        page_h = (section.page_height or 10692130) / EMU_PER_PT
        margins = (
            (section.left_margin or 914400) / EMU_PER_PT,
            (section.right_margin or 914400) / EMU_PER_PT,
            (section.top_margin or 914400) / EMU_PER_PT,
            (section.bottom_margin or 914400) / EMU_PER_PT,
        )
    else:
        page_w, page_h, margins = 595.3, 841.9, (72, 72, 72, 72)

    try:
        base_size = float(doc.styles["Normal"].font.size.pt)  # type: ignore[union-attr]
    except Exception:
        base_size = 10.5
    base_size = base_size if 6 <= base_size <= 24 else 10.5

    writer = _Writer(doc, page_w, page_h, margins)
    empty_run = 0

    for block in _iter_blocks(doc):
        if isinstance(block, Table):
            writer.table(block)
            empty_run = 0
            continue

        para = block
        style_name = (para.style.name or "").strip()
        level = _heading_level(para)
        text = para.text.strip() if para.text else ""
        has_image = bool(para._p.findall(".//" + qn("a:blip")))

        if not text and not has_image:
            empty_run += 1
            if empty_run <= 1:
                writer.paragraph("", size=base_size, space_after=2)
            continue
        empty_run = 0

        if style_name.lower() in _TITLE_STYLES:
            writer.paragraph(
                text, size=22.0, bold=True, align=_alignment(para),
                space_before=6, space_after=14,
            )
            continue

        if level:
            size = _HEADING_SIZE.get(level, base_size + 1)
            writer.paragraph(
                text, size=size, bold=True, align=_alignment(para),
                space_before=size * 0.55 if writer.y > margins[2] + 1 else 0,
                space_after=size * 0.35,
            )
            continue

        if _is_list_item(para):
            text = "• " + text
        writer.paragraph(
            text,
            size=base_size,
            bold=any(r.bold for r in para.runs if r.bold),
            align=_alignment(para),
        )
        if has_image:
            for blob, w_pt, h_pt in _para_images(para):
                writer.image(blob, w_pt, h_pt)

    if writer.out.page_count == 0:
        writer._new_page()
    try:
        writer.save(dst)
    except Exception as exc:
        raise ConversionError(f"写 PDF 失败：{exc}") from exc
