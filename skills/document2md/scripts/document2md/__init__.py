"""Format-aware PDF and Microsoft Office to Markdown conversion."""

from .converter import ConversionOptions, ConversionResult, convert_file, convert_files

__all__ = ["ConversionOptions", "ConversionResult", "convert_file", "convert_files"]
__version__ = "0.1.0"
