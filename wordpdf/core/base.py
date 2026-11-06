"""引擎基类。每个引擎只需回答两件事：能不能用？怎么转？"""

from __future__ import annotations

from pathlib import Path

from .model import ConversionError, Options


class BaseEngine:
    name: str = "base"
    label: str = "未命名引擎"
    note: str = ""
    #: 该引擎支持的源格式后缀；None 表示由引擎自己判断
    suffixes: set[str] | None = None

    def probe(self) -> tuple[bool, str]:
        """返回 (是否可用, 说明)。不抛异常。"""
        return False, "未实现"

    def convert(self, src: Path, dst: Path, opts: Options) -> None:
        """把 src 转成 dst。失败请抛 ConversionError。"""
        raise ConversionError("未实现")

    # -- 便捷方法 ---------------------------------------------------------
    def supports(self, suffix: str) -> bool:
        return self.suffixes is None or suffix.lower() in self.suffixes

    def __repr__(self) -> str:  # pragma: no cover
        ok, msg = self.probe()
        return f"<{self.name} available={ok} {msg}>"
