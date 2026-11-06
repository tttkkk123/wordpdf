"""命令行入口：python -m wordpdf ..."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from . import __version__
from .core import convert_many, default_target, detect_direction, engine_table
from .core.discovery import describe_environment
from .core.model import ConversionError, Options


def _build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="wordpdf",
        description="Word ↔ PDF 双向转换（混合引擎，自动降级）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "示例:\n"
            "  wordpdf 报告.docx                    # Word → PDF（同目录）\n"
            "  wordpdf 合同.pdf -o D:\\out           # PDF → Word\n"
            "  wordpdf *.docx --engine word-com     # 指定引擎\n"
            "  wordpdf 手册.pdf --pages 1-10        # 只转前 10 页\n"
            "  wordpdf --info                       # 体检本机环境与引擎\n"
        ),
    )
    p.add_argument("files", nargs="*", help="待转换文件（可多个）")
    p.add_argument("-v", "--version", action="version", version=f"wordpdf {__version__}")
    p.add_argument("--info", action="store_true", help="显示环境体检与各引擎可用性")
    p.add_argument("--to", choices=("auto", "pdf", "docx"), default="auto",
                   help="目标方向，默认按后缀自动判断")
    p.add_argument("-o", "--outdir", help="输出目录，默认与源文件同目录")
    p.add_argument("--engine", help="强制指定引擎（word-com / libreoffice / python-render / pdf2docx / text-flow）")
    p.add_argument("--pages", default="", help="PDF→Word 的页码范围，如 1-10")
    p.add_argument("--screen", action="store_true", help="Word→PDF 按屏幕质量导出（文件更小）")
    p.add_argument("--pdfa", action="store_true", help="Word→PDF 输出 PDF/A-1b 归档格式")
    p.add_argument("--no-bookmarks", action="store_true", help="不生成 PDF 书签")
    p.add_argument("--overwrite", action="store_true", help="允许覆盖同名输出文件")
    p.add_argument("-q", "--quiet", action="store_true", help="静默模式，只输出结果")
    return p


def _pad(text: str, width: int) -> str:
    """按显示宽度补空格（中文占两列），保证终端里表格对齐。"""
    import unicodedata

    shown = sum(2 if unicodedata.east_asian_width(ch) in "WF" else 1 for ch in text)
    return text + " " * max(0, width - shown)


def _print_info() -> None:
    print("运行环境")
    print("-" * 70)
    for name, msg, ok in describe_environment():
        print(f"  [{'√' if ok else '×'}] {_pad(name, 18)}{msg}")
    for direction, title in (("docx2pdf", "Word → PDF"), ("pdf2docx", "PDF → Word")):
        print(f"\n{title} 引擎链（自上而下依次尝试）")
        print("-" * 70)
        for i, (name, label, ok, msg) in enumerate(engine_table(direction), 1):
            state = "可用" if ok else "不可用"
            print(f"  {i}. {_pad(f'[{state}]', 9)}{_pad(label, 30)}{msg}")
    print()


def main(argv: list[str] | None = None) -> int:
    args = _build_parser().parse_args(argv)

    if args.info:
        _print_info()
        if not args.files:
            return 0

    if not args.files:
        _build_parser().print_help()
        return 1

    direction = {"auto": "auto", "pdf": "docx2pdf", "docx": "pdf2docx"}[args.to]

    opts = Options(
        overwrite=args.overwrite,
        optimize_for=1 if args.screen else 0,
        pdfa=args.pdfa,
        bookmarks=not args.no_bookmarks,
        pages=args.pages,
    )

    files = []
    for pattern in args.files:
        p = Path(pattern)
        if p.is_file():
            files.append(p)
        else:
            hits = sorted(Path().glob(pattern))
            if not hits:
                print(f"跳过：找不到 {pattern}", file=sys.stderr)
            files.extend(h for h in hits if h.is_file())
    if not files:
        print("没有可转换的文件", file=sys.stderr)
        return 1

    def log(msg: str) -> None:
        if not args.quiet:
            print(msg)

    results = convert_many(
        files,
        outdir=args.outdir,
        direction=direction,  # type: ignore[arg-type]
        options=opts,
        engine=args.engine,
        log=log,
    )

    ok = sum(1 for r in results if r.ok)
    print(f"\n完成：成功 {ok} / 共 {len(results)}")
    for r in results:
        if not r.ok:
            print(f"  ✗ {r.src.name}  {r.message}")
    return 0 if ok == len(results) else 2


if __name__ == "__main__":
    raise SystemExit(main())
