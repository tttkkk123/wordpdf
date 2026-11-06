"""转换编排层：选引擎、定路径、失败了就沿降级链往下走。"""

from __future__ import annotations

import time
from pathlib import Path
from typing import Callable, Iterable, Optional

from .docx2pdf import LibreOfficeEngine, PythonRendererEngine, WordComEngine
from .model import (
    PDF_SUFFIXES,
    WORD_SUFFIXES,
    ConversionError,
    ConversionResult,
    Direction,
    Options,
)
from .pdf2docx_engines import Pdf2DocxEngine, TextFlowEngine

Logger = Callable[[str], None]

#: 降级链：顺序即优先级
ENGINES: dict[str, list] = {
    "docx2pdf": [WordComEngine(), LibreOfficeEngine(), PythonRendererEngine()],
    "pdf2docx": [Pdf2DocxEngine(), TextFlowEngine()],
}

DIRECTION_BY_SUFFIX = {
    **{s: "docx2pdf" for s in WORD_SUFFIXES},
    **{s: "pdf2docx" for s in PDF_SUFFIXES},
}


def engine_table(direction: str) -> list[tuple[str, str, bool, str]]:
    """[(name, label, 可用, 说明)]，供 CLI/GUI 展示。"""
    rows = []
    for eng in ENGINES.get(direction, []):
        ok, msg = eng.probe()
        rows.append((eng.name, eng.label, ok, msg))
    return rows


def detect_direction(src: Path, wanted: Direction = "auto") -> str:
    if wanted != "auto":
        return wanted
    suffix = src.suffix.lower()
    if suffix in PDF_SUFFIXES:
        return "pdf2docx"
    if suffix in WORD_SUFFIXES:
        return "docx2pdf"
    raise ConversionError(f"不认识的文件类型：{suffix or '(无后缀)'}")


def default_target(src: Path, direction: str, outdir: Optional[Path] = None,
                   overwrite: bool = False) -> Path:
    suffix = ".pdf" if direction == "docx2pdf" else ".docx"
    folder = Path(outdir) if outdir else src.parent
    candidate = folder / f"{src.stem}{suffix}"
    if overwrite or not candidate.exists():
        return candidate
    idx = 1
    while True:
        candidate = folder / f"{src.stem}_{idx}{suffix}"
        if not candidate.exists():
            return candidate
        idx += 1


def convert_file(
    src: Path | str,
    dst: Path | str | None = None,
    direction: Direction = "auto",
    options: Optional[Options] = None,
    engine: Optional[str] = None,
    log: Optional[Logger] = None,
) -> ConversionResult:
    """转换单个文件。不抛异常，失败信息放在 result 里。"""
    opts = options or Options()
    src = Path(src).expanduser().resolve()
    say = log or (lambda _m: None)

    if not src.exists():
        return ConversionResult(ok=False, src=src, message="文件不存在")
    if src.stat().st_size == 0:
        return ConversionResult(ok=False, src=src, message="文件为空")

    try:
        real_direction = detect_direction(src, direction)
    except ConversionError as exc:
        return ConversionResult(ok=False, src=src, message=str(exc))

    target = Path(dst).expanduser() if dst else default_target(
        src, real_direction,
        overwrite=opts.overwrite,
    )
    target = target.resolve()
    target.parent.mkdir(parents=True, exist_ok=True)
    if target.exists() and not opts.overwrite and dst is not None:
        return ConversionResult(ok=False, src=src, dst=target, message="目标文件已存在（未开启覆盖）")

    chain = ENGINES[real_direction]
    if engine:
        chain = [e for e in chain if e.name == engine] or []
        if not chain:
            return ConversionResult(ok=False, src=src, message=f"未知引擎：{engine}")

    attempts: list[tuple[str, str]] = []
    result = ConversionResult(ok=False, src=src, dst=target, direction=real_direction)

    for idx, eng in enumerate(chain):
        ok, msg = eng.probe()
        if not ok:
            attempts.append((eng.name, f"不可用：{msg}"))
            say(f"  跳过 {eng.label} —— {msg}")
            continue
        say(f"  使用引擎：{eng.label}")
        started = time.perf_counter()
        try:
            eng.convert(src, target, opts)
            elapsed = time.perf_counter() - started
            if not target.exists() or target.stat().st_size == 0:
                raise ConversionError("引擎执行完毕但没有产出文件")
            result.ok = True
            result.engine = eng.name
            result.engine_label = eng.label
            result.elapsed = elapsed
            result.size = target.stat().st_size
            result.fallback = bool(attempts)  # 之前有过跳过或失败，即视为降级
            result.attempts = attempts
            return result
        except Exception as exc:  # 任何异常都进入降级
            elapsed = time.perf_counter() - started
            detail = str(exc) or exc.__class__.__name__
            attempts.append((eng.name, detail))
            say(f"  引擎失败（{elapsed:.2f}s）：{detail}")
            if target.exists() and target.stat().st_size == 0:
                try:
                    target.unlink()
                except OSError:
                    pass

    result.attempts = attempts
    result.message = "；".join(f"{n}: {m}" for n, m in attempts) or "没有可用引擎"
    return result


def convert_many(
    files: Iterable[Path | str],
    outdir: Optional[Path | str] = None,
    direction: Direction = "auto",
    options: Optional[Options] = None,
    engine: Optional[str] = None,
    log: Optional[Logger] = None,
    on_file: Optional[Callable[[int, int, ConversionResult], None]] = None,
) -> list[ConversionResult]:
    opts = options or Options()
    files = [Path(f) for f in files]
    out = Path(outdir).expanduser().resolve() if outdir else None
    results: list[ConversionResult] = []
    say = log or (lambda _m: None)

    for i, src in enumerate(files, 1):
        say(f"[{i}/{len(files)}] {src.name}")
        try:
            d = detect_direction(src, direction)
        except ConversionError as exc:
            res = ConversionResult(ok=False, src=src, message=str(exc))
            results.append(res)
            if on_file:
                on_file(i, len(files), res)
            continue
        dst = default_target(src.resolve(), d, outdir=out, overwrite=opts.overwrite)
        res = convert_file(src, dst, direction=d, options=opts, engine=engine, log=say)
        results.append(res)
        if on_file:
            on_file(i, len(files), res)   # 界面自己会展示逐条结果
        else:
            say("  " + res.summary)
    return results
