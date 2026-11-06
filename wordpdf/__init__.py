"""wordpdf —— Word / PDF 双向转换插件。

实现方式：混合引擎（Hybrid Engine）+ 自动降级链。
    Word -> PDF : Word COM 自动化(保真) -> LibreOffice 无头模式 -> 纯 Python 兜底排版
    PDF  -> Word: pdf2docx 版面重建 -> PyMuPDF 纯文本流重建
"""

__version__ = "1.0.1"
__all__ = ["__version__"]
