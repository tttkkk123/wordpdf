"""核心层：引擎探测、双向转换、降级编排。"""

from .converter import (
    ENGINES,
    convert_file,
    convert_many,
    default_target,
    detect_direction,
    engine_table,
)
from .model import ConversionError, ConversionResult, Direction, Options

__all__ = [
    "ENGINES",
    "ConversionError",
    "ConversionResult",
    "Direction",
    "Options",
    "convert_file",
    "convert_many",
    "default_target",
    "detect_direction",
    "engine_table",
]
