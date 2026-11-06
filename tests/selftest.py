"""自检脚本：造样例 → 双向转换 → 校验产物。

用法：
    .venv\\Scripts\\python.exe tests\\selftest.py
"""

from __future__ import annotations

import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from docx import Document                       # noqa: E402
from docx.enum.text import WD_ALIGN_PARAGRAPH  # noqa: E402
from docx.shared import Pt, RGBColor            # noqa: E402

from wordpdf.core import convert_file, engine_table  # noqa: E402
from wordpdf.core.discovery import describe_environment  # noqa: E402
from wordpdf.core.model import Options  # noqa: E402

OUT = ROOT / "tests" / "_out"


def make_sample(path: Path) -> Path:
    """造一份带标题、正文、列表、表格、图片（矢量图）的 docx。"""
    doc = Document()
    doc.styles["Normal"].font.name = "Microsoft YaHei"
    doc.styles["Normal"].font.size = Pt(10.5)

    title = doc.add_heading("季度经营分析报告", level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p = doc.add_paragraph("编制部门：战略与运营部    报告期：2026 年第三季度")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in p.runs:
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0x66, 0x66, 0x66)

    doc.add_heading("一、整体经营情况", level=1)
    doc.add_paragraph(
        "本季度公司实现营业收入 12.8 亿元，同比增长 18.4%；归母净利润 1.62 亿元，"
        "同比增长 9.7%。毛利率 31.2%，环比提升 0.8 个百分点，主要得益于高毛利产品占比提升，"
        "以及供应链集中采购带来的成本优化。"
    )
    doc.add_heading("二、重点事项", level=2)
    for item in (
        "华东区新增 3 个区域仓，平均履约时效缩短至 11.6 小时。",
        "研发费用投入 0.94 亿元，占营收 7.3%，同比增加 1.1 个百分点。",
        "完成对上游两家供应商的整合，采购成本下降约 4.2%。",
    ):
        doc.add_paragraph(item, style="List Bullet")

    doc.add_heading("三、分区域经营数据", level=2)
    rows = [
        ("区域", "营收（亿元）", "同比", "毛利率"),
        ("华东", "4.86", "+21.3%", "33.5%"),
        ("华南", "3.12", "+17.8%", "31.0%"),
        ("华北", "2.94", "+12.4%", "29.2%"),
        ("西部", "1.88", "+9.1%", "27.6%"),
    ]
    table = doc.add_table(rows=len(rows), cols=4)
    table.style = "Table Grid"
    for r, row in enumerate(rows):
        for c, value in enumerate(row):
            cell = table.cell(r, c)
            cell.text = value
            if r == 0:
                for run in cell.paragraphs[0].runs:
                    run.bold = True

    doc.add_heading("四、下一步工作", level=1)
    doc.add_paragraph(
        "四季度将围绕「提效、稳价、拓新」三条主线推进：一是推进仓配一体化，把履约成本再压降 3%；"
        "二是稳住核心品类价格体系，避免无效促销；三是启动两个新品类孵化，力争年底前完成首批铺货。"
    )
    doc.save(str(path))
    return path


def make_pdf_sample(path: Path) -> Path:
    """再单独造一份 PDF（用 PyMuPDF 画），用来测 PDF→Word 链路。"""
    import pymupdf

    doc = pymupdf.open()
    page = doc.new_page(width=595, height=842)
    y = 90
    page.insert_text((72, y), "采购合同摘要", fontsize=20, fontname="china-s")
    y += 40
    body = (
        "甲方：某某科技有限公司        乙方：某某供应链管理有限公司\n"
        "合同金额：人民币 286 万元（含税）        期限：2026-01-01 至 2026-12-31\n"
        "结算方式：月度对账，次月 15 日前付款。逾期按日万分之五计收违约金。"
    )
    for line in body.split("\n"):
        page.insert_text((72, y), line, fontsize=10.5, fontname="china-s")
        y += 20
    y += 10
    page.insert_text((72, y), "第一条  交付与验收", fontsize=13, fontname="china-s")
    y += 24
    page.insert_text(
        (72, y),
        "乙方应按甲方订单约定的时间、数量与质量标准交付货物。甲方在收货后 7 个工作日内完成验收，",
        fontsize=10.5, fontname="china-s",
    )
    y += 19
    page.insert_text(
        (72, y),
        "验收不合格的，乙方应在 5 个工作日内完成更换或退货，并承担由此产生的全部费用。",
        fontsize=10.5, fontname="china-s",
    )
    doc.save(str(path))
    doc.close()
    return path


def head(text: str) -> None:
    print("\n" + "=" * 72)
    print(text)
    print("=" * 72)


def main() -> int:
    OUT.mkdir(parents=True, exist_ok=True)
    failures: list[str] = []

    head("1. 环境体检")
    for name, msg, ok in describe_environment():
        print(f"  [{'√' if ok else '×'}] {name:<16} {msg}")
    for d, title in (("docx2pdf", "Word → PDF"), ("pdf2docx", "PDF → Word")):
        print(f"  {title} 引擎链: " + " → ".join(
            f"{lb}{'' if ok else '(不可用)'}" for _n, lb, ok, _m in engine_table(d)
        ))

    head("2. 生成样例文件")
    sample_docx = make_sample(OUT / "样例报告.docx")
    sample_pdf = make_pdf_sample(OUT / "样例合同.pdf")
    print(f"  {sample_docx.name}  {sample_docx.stat().st_size} 字节")
    print(f"  {sample_pdf.name}  {sample_pdf.stat().st_size} 字节")

    head("3. Word → PDF（默认引擎链）")
    r1 = convert_file(sample_docx, OUT / "样例报告.pdf", options=Options(overwrite=True), log=print)
    print("  " + r1.summary)
    if not r1.ok:
        failures.append(f"docx→pdf: {r1.message}")
    else:
        import pymupdf

        with pymupdf.open(str(r1.dst)) as pdf:
            print(f"  页数={pdf.page_count}  文字层字符数={len(pdf[0].get_text())}")
            if "季度经营分析报告" not in pdf[0].get_text():
                failures.append("docx→pdf: 首页文字缺失")

    head("4. PDF → Word（pdf2docx 版面重建）")
    if r1.ok:
        r2 = convert_file(r1.dst, OUT / "样例报告_回程.docx", options=Options(overwrite=True), log=print)
        print("  " + r2.summary)
        if not r2.ok:
            failures.append(f"pdf→docx: {r2.message}")
        else:
            back = Document(str(r2.dst))
            text = "\n".join(p.text for p in back.paragraphs)
            print(f"  段落={len(back.paragraphs)}  表格={len(back.tables)}  字符={len(text)}")
            if "季度经营分析报告" not in text:
                failures.append("pdf→docx: 标题未回填")
            if len(back.tables) == 0:
                print("  提示：未识别出表格（pdf2docx 对无边框表格识别较弱）")

    head("5. 纯 Python 兜底引擎（不依赖 Office）")
    t0 = time.perf_counter()
    r3 = convert_file(sample_docx, OUT / "样例报告_pure.pdf", engine="python-render", options=Options(overwrite=True), log=print)
    print(f"  {r3.summary}  (总耗时 {time.perf_counter() - t0:.2f}s)")
    if not r3.ok:
        failures.append(f"python-render: {r3.message}")
    else:
        import pymupdf

        with pymupdf.open(str(r3.dst)) as pdf:
            print(f"  页数={pdf.page_count}")
            txt = "".join(p.get_text() for p in pdf)
            if "季度经营分析报告" not in txt:
                failures.append("python-render: 标题未渲染")

    head("6. PDF → Word 纯文本流兜底")
    r4 = convert_file(sample_pdf, OUT / "样例合同_textflow.docx", engine="text-flow", options=Options(overwrite=True), log=print)
    print("  " + r4.summary)
    if not r4.ok:
        failures.append(f"text-flow: {r4.message}")

    head("7. 结果")
    print(f"  产物目录：{OUT}")
    for f in sorted(OUT.glob("*")):
        print(f"    {f.name:<28} {f.stat().st_size:>9} 字节")
    if failures:
        print("\n  ✗ 失败项：")
        for f in failures:
            print(f"    - {f}")
        return 1
    print("\n  √ 全部通过")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
