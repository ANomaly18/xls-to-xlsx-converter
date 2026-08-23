"""XLS to XLSX Converter — offline conversion library and desktop tool."""

__version__ = "1.0.0"

from .converter import (
    KNOWN_FILE_TYPES,
    convert_file,
    decode_with_fallback,
    detect_file_format,
    identify_and_convert,
)

__all__ = [
    "KNOWN_FILE_TYPES",
    "convert_file",
    "decode_with_fallback",
    "detect_file_format",
    "identify_and_convert",
    "__version__",
]
