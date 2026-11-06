"""公共数据模型：转换方向、选项、结果。"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Literal, Optional

Direction = Literal["auto", "docx2pdf", "pdf2docx"]

# 可被 Word 打开并导出为 PDF 的源格式
WORD_SUFFIXES = {".docx", ".doc", ".docm", ".dotx", ".dot", ".rtf", ".odt", ".wps"}
PDF_SUFFIXES = {".pdf"}

DIRECTION_LABEL = {
    "docx2pdf": "Word → PDF",
    "pdf2docx": "PDF → Word",
    "auto": "自动识别",
}


@dataclass
class Options:
    """各引擎共用的转换选项（不支持的选项会被静默忽略）。"""

    overwrite: bool = False
    # 仅 Word -> PDF
    optimize_for: int = 0        # 0 = 按打印质量, 1 = 按屏幕质量
    pdfa: bool = False           # 输出 PDF/A-1b 归档格式
    bookmarks: bool = True       # 用标题生成 PDF 书签
    # 仅 PDF -> Word
    pages: str = ""              # 页码范围，如 "1-5" 或 "3"，空 = 全部
    multi_processing: bool = False  # pdf2docx 多进程（页数很多时更快）

    def page_range(self) -> tuple[int, Optional[int]]:
        """把 "2-7" 解析成 pdf2docx 需要的 (start, end)，end 为开区间。"""
        text = (self.pages or "").strip().replace("，", ",")
        if not text:
            return 0, None
        head = text.split(",")[0]
        if "-" in head:
            a, b = head.split("-", 1)
            try:
                start = max(0, int(a.strip()) - 1)
                end = int(b.strip())
                return start, end
            except ValueError:
                return 0, None
        try:
            n = int(head) - 1
            return max(0, n), n + 1
        except ValueError:
            return 0, None


@dataclass
class ConversionResult:
    ok: bool
    src: Path
    dst: Optional[Path] = None
    direction: str = ""
    engine: str = ""          # 实际生效的引擎 name
    engine_label: str = ""
    elapsed: float = 0.0
    size: int = 0
    message: str = ""
    fallback: bool = False    # 是否发生过降级
    attempts: list[tuple[str, str]] = field(default_factory=list)  # [(engine, 失败原因)]

    @property
    def summary(self) -> str:
        if not self.ok:
            return f"失败  {self.src.name}：{self.message}"
        flag = "（已降级）" if self.fallback else ""
        return (
            f"成功  {self.src.name} → {self.dst.name if self.dst else '-'} "
            f"[{self.engine_label or self.engine}{flag}] {self.elapsed:.2f}s"
        )


class ConversionError(RuntimeError):
    """带用户可读信息的转换异常。"""
