"""字体解析。

PyMuPDF 内置的 "china-s" 会把 ASCII 当成全角渲染（数字被拉开），
所以优先嵌入系统 TrueType 字体；实在找不到才退回内置字体。
"""

from __future__ import annotations

import os
from dataclasses import dataclass
from functools import lru_cache

try:
    import pymupdf as fitz  # type: ignore
except ImportError:  # pragma: no cover
    import fitz  # type: ignore

_REGULAR = [
    r"C:\Windows\Fonts\Deng.ttf",          # 等线：现代无衬线，空格映射正确
    r"C:\Windows\Fonts\simsun.ttc",        # 宋体
    r"C:\Windows\Fonts\msyh.ttc",          # 微软雅黑（TTC 会把空格映射成 nbsp，排在后面）
    "/System/Library/Fonts/PingFang.ttc",
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
    "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
]
_BOLD = [
    r"C:\Windows\Fonts\Dengb.ttf",
    r"C:\Windows\Fonts\simhei.ttf",
    r"C:\Windows\Fonts\msyhbd.ttc",
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc",
]


@dataclass(frozen=True)
class FontSpec:
    alias: str
    file: str | None
    font: object | None

    def width(self, text: str, size: float) -> float:
        if self.font is not None:
            try:
                return float(self.font.text_length(text, fontsize=size))  # type: ignore[attr-defined]
            except Exception:
                pass
        try:
            return float(fitz.get_text_length(text, fontname=self.alias, fontsize=size))
        except Exception:
            return sum(size if ord(ch) > 0x2E7F else size * 0.52 for ch in text)


@lru_cache(maxsize=4)
def _resolve(kind: str) -> FontSpec:
    candidates = _BOLD if kind == "bold" else _REGULAR
    alias = "CJKBold" if kind == "bold" else "CJKReg"
    for path in candidates:
        if not os.path.exists(path):
            continue
        try:
            return FontSpec(alias=alias, file=path, font=fitz.Font(fontfile=path))
        except Exception:
            continue
    return FontSpec(alias="china-s", file=None, font=None)


def regular() -> FontSpec:
    return _resolve("regular")


def bold() -> FontSpec:
    """没有粗体字体文件时返回常规字体，由调用方用重绘模拟加粗。"""
    return _resolve("bold")
