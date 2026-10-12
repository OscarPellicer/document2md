"""PDF conversion with Datalab's Chandra OCR model.

Chandra is a vision-language model: every page is rendered to an image and
transcribed into layout-tagged HTML, which is then turned into Markdown. It
reads scans and born-digital PDFs the same way, writes equations as LaTeX and
citation markers as ``<sup>``, and returns a crop for each figure.

It runs either locally through Hugging Face ``transformers`` (CUDA, Apple MPS
or CPU) or against an OpenAI-compatible server that hosts the model (vLLM,
``mlx_vlm.server`` on Apple Silicon, llama.cpp). The ``chandra-ocr`` package is
an optional dependency, installed on first use by the entry point.
"""

from __future__ import annotations

import base64
import os
import re
from dataclasses import dataclass, field
from io import BytesIO
from pathlib import Path
from urllib.parse import quote


SERVER_ENV = "DOCUMENT2MD_CHANDRA_SERVER"
MODEL_ENV = "DOCUMENT2MD_CHANDRA_MODEL"
INSTALL_HINT = (
    "The Chandra PDF backend needs the optional `chandra-ocr` package. Run "
    "convert_document.py without --no-bootstrap so it installs "
    "scripts/requirements-chandra.txt into the skill environment, or install "
    "that file yourself."
)
# Chandra renders pages at 192 dpi, and at least 1024 px on the short side.
IMAGE_DPI = 192
MIN_PAGE_PIXELS = 1024
IMAGE_PLACEHOLDER = "<!-- image -->"
_MARKDOWN_IMAGE = re.compile(r"!\[([^\]]*)\]\(([^)\s]+)\)")
_HTML_TABLE = re.compile(r"<table\b.*?</table>", re.S | re.I)


@dataclass(slots=True)
class ChandraEngine:
    """A loaded Chandra model, or a client for a server that hosts one."""

    manager: object
    method: str
    server: str | None = None
    generate_kwargs: dict = field(default_factory=dict)

    @property
    def description(self) -> str:
        return f"server {self.server}" if self.server else "local transformers"


def create_engine(server: str | None = None, model: str | None = None) -> ChandraEngine:
    """Load Chandra locally, or prepare a client for ``server``.

    ``server`` and ``model`` default to the ``DOCUMENT2MD_CHANDRA_SERVER`` and
    ``DOCUMENT2MD_CHANDRA_MODEL`` environment variables.
    """
    server = server or os.environ.get(SERVER_ENV) or None
    model = model or os.environ.get(MODEL_ENV) or None
    try:
        from chandra.model import InferenceManager
        from chandra.settings import settings
    except ImportError as error:
        raise RuntimeError(INSTALL_HINT) from error

    if server:
        # Chandra reads the served model name from its settings. None makes it
        # ask the server, which suits servers that host one model.
        settings.VLLM_MODEL_NAME = model or os.environ.get("VLLM_MODEL_NAME") or None
        return ChandraEngine(
            InferenceManager(method="vllm"),
            "vllm",
            server,
            {"vllm_api_base": server.rstrip("/"), "max_workers": 1},
        )

    if model:
        settings.MODEL_CHECKPOINT = model
    if not settings.TORCH_DEVICE:
        settings.TORCH_DEVICE = _torch_device()
    try:
        manager = InferenceManager(method="hf")
    except ImportError as error:
        raise RuntimeError(INSTALL_HINT) from error
    return ChandraEngine(manager, "hf")


def _torch_device() -> str | None:
    try:
        import torch
    except ImportError:
        return None
    if torch.cuda.is_available():
        return None  # let accelerate spread the model over the GPUs
    if getattr(torch.backends, "mps", None) and torch.backends.mps.is_available():
        return "mps"
    return "cpu"


def _render_pages(source: bytes, first: int, last: int):
    """Render 1-based pages ``first``..``last`` the way Chandra's own loader does."""
    import pypdfium2 as pdfium
    import pypdfium2.raw as pdfium_c

    document = pdfium.PdfDocument(source)
    try:
        document.init_forms()
        images = []
        for index in range(first - 1, last):
            page = document[index]
            pdfium_c.FPDFPage_Flatten(page, pdfium_c.FLAT_NORMALDISPLAY)
            page = document[index]
            short_side = min(page.get_width(), page.get_height())
            dpi = max(IMAGE_DPI, MIN_PAGE_PIXELS / short_side * 72)
            images.append(page.render(scale=dpi / 72).to_pil().convert("RGB"))
        return images
    finally:
        document.close()


def _cell_markdown(cell, converter) -> str:
    text = converter.convert(cell.decode_contents()).strip()
    text = re.sub(r"\s*\n\s*", "<br>", text)
    return text.replace("|", "\\|")


def _html_table_to_markdown(html: str, converter) -> str:
    """Lay an HTML table out on its grid; spanned cells appear once, other positions stay blank."""
    from bs4 import BeautifulSoup

    table = BeautifulSoup(html, "html.parser").find("table")
    if table is None:
        return html
    grid: dict[tuple[int, int], str] = {}
    caption = table.find("caption")
    rows = table.find_all("tr")
    for row_index, row in enumerate(rows):
        column = 0
        for cell in row.find_all(["td", "th"], recursive=False):
            while (row_index, column) in grid:
                column += 1
            try:
                rowspan = max(1, int(cell.get("rowspan", 1)))
                colspan = max(1, int(cell.get("colspan", 1)))
            except ValueError:
                rowspan = colspan = 1
            for r in range(row_index, row_index + rowspan):
                for c in range(column, column + colspan):
                    grid[(r, c)] = ""
            grid[(row_index, column)] = _cell_markdown(cell, converter)
            column += colspan
    if not grid:
        return ""
    height = max(r for r, _ in grid) + 1
    width = max(c for _, c in grid) + 1
    lines = []
    for r in range(height):
        lines.append("| " + " | ".join(grid.get((r, c), "") for c in range(width)) + " |")
        if r == 0:
            lines.append("| " + " | ".join("---" for _ in range(width)) + " |")
    markdown = "\n".join(lines)
    if caption is not None:
        markdown = converter.convert(caption.decode_contents()).strip() + "\n\n" + markdown
    return markdown


def _markdown_converter():
    from chandra.output import Markdownify

    return Markdownify(
        heading_style="ATX",
        bullets="-",
        escape_misc=False,
        escape_underscores=True,
        escape_asterisks=True,
        escape_dollars=True,
        sub_symbol="<sub>",
        sup_symbol="<sup>",
        inline_math_delimiters=("$", "$"),
        block_math_delimiters=("$$", "$$"),
    )


def _tables_to_markdown(markdown: str, converter) -> str:
    return _HTML_TABLE.sub(lambda match: _html_table_to_markdown(match.group(0), converter), markdown)


def _image_bytes(image, fmt: str = "PNG") -> bytes:
    buffer = BytesIO()
    image.save(buffer, format=fmt)
    return buffer.getvalue()


def _place_images(
    markdown: str,
    images: dict,
    page_number: int,
    mode: str,
    output_path: Path,
    assets_dir: Path,
    assets: list[Path],
) -> str:
    """Replace Chandra's image references according to the requested image mode."""
    counter = 0

    def replace(match: re.Match[str]) -> str:
        nonlocal counter
        alt, target = match.group(1), match.group(2)
        image = images.get(target)
        if image is None or mode == "placeholder":
            return IMAGE_PLACEHOLDER
        counter += 1
        if mode == "embed":
            encoded = base64.b64encode(_image_bytes(image)).decode("ascii")
            return f"![{alt}](data:image/png;base64,{encoded})"
        assets_dir.mkdir(parents=True, exist_ok=True)
        path = assets_dir / f"page-{page_number:04d}-image-{counter}.png"
        path.write_bytes(_image_bytes(image))
        assets.append(path)
        relative = Path(os.path.relpath(path, output_path.parent)).as_posix()
        return f"![{alt}]({quote(relative)})"

    return _MARKDOWN_IMAGE.sub(replace, markdown)


def convert_pdf(
    engine: ChandraEngine,
    path: Path,
    output_path: Path,
    images: str,
    assets_dir: Path,
    page_ranges: list[tuple[int, int]],
) -> tuple[str, list[Path], list[str]]:
    """Transcribe ``path`` page by page and return Markdown, assets and warnings."""
    from chandra.model.schema import BatchInputItem
    from chandra.model.util import detect_repeat_token

    source = path.read_bytes()
    converter = _markdown_converter()
    parts: list[str] = []
    assets: list[Path] = []
    warnings: list[str] = []

    for first, last in page_ranges:
        pages = _render_pages(source, first, last)
        # Pages go one at a time: on a single GPU or server batching saves
        # little, and one page's failure stays its own.
        for offset, image in enumerate(pages):
            page_number = first + offset
            (result,) = engine.manager.generate(
                [BatchInputItem(image=image, prompt_type="ocr_layout")],
                include_images=True,
                include_headers_footers=False,
                **engine.generate_kwargs,
            )
            if result.error or not result.raw.strip():
                warnings.append(f"Page {page_number}: Chandra returned no text.")
                parts.append(f"<!-- page {page_number}: not transcribed -->")
                continue
            if detect_repeat_token(result.raw):
                warnings.append(
                    f"Page {page_number}: Chandra output ends in a repetition loop; check this page."
                )
            markdown = _tables_to_markdown(result.markdown, converter)
            markdown = _place_images(
                markdown, result.images, page_number, images, output_path, assets_dir, assets
            )
            parts.append(markdown.strip())
        del pages

    markdown = "\n\n".join(part for part in parts if part)
    return markdown.rstrip() + "\n", assets, warnings
