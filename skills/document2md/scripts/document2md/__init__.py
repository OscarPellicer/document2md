"""Format-aware PDF, Office and OpenDocument to Markdown conversion."""

from .converter import (
    BatchResult,
    ConversionFailure,
    ConversionOptions,
    ConversionResult,
    convert_file,
    convert_files,
)

__all__ = [
    "BatchResult",
    "ConversionFailure",
    "ConversionOptions",
    "ConversionResult",
    "convert_file",
    "convert_files",
]
__version__ = "0.2.0"
