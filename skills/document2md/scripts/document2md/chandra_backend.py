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
PICTURE_LABELS = {"Image", "Figure", "Diagram"}
SKIPPED_LABELS = {"Blank-Page", "Page-Header", "Page-Footer"}
_DOUBLED_BULLET = re.compile(r"^(\s*[-*+] )[\u2022\u00b7\u25aa\u25cf\u25e6\u2023\u2013]\s*", re.M)
_DOUBLED_NUMBER = re.compile(r"^(\s*)(\d{1,3})\. \2[.)]\s+", re.M)
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


def _image_bytes(image) -> bytes:
    buffer = BytesIO()
    image.save(buffer, format="PNG")
    return buffer.getvalue()


def _picture_markdown(
    page_image,
    bbox: list[int],
    page_number: int,
    index: int,
    mode: str,
    output_path: Path,
    assets_dir: Path,
    assets: list[Path],
) -> str:
    """Crop one figure from the rendered page and reference it per image mode."""
    left, top, right, bottom = bbox
    if mode == "placeholder" or right - left < 2 or bottom - top < 2:
        return IMAGE_PLACEHOLDER
    crop = page_image.crop((left, top, right, bottom))
    if mode == "embed":
        encoded = base64.b64encode(_image_bytes(crop)).decode("ascii")
        return f"![Image](data:image/png;base64,{encoded})"
    assets_dir.mkdir(parents=True, exist_ok=True)
    path = assets_dir / f"page-{page_number:04d}-image-{index}.png"
    path.write_bytes(_image_bytes(crop))
    assets.append(path)
    relative = Path(os.path.relpath(path, output_path.parent)).as_posix()
    return f"![Image]({quote(relative)})"


def _text_block_markdown(label: str, html: str, converter) -> str:
    from bs4 import BeautifulSoup

    soup = BeautifulSoup(html, "html.parser")
    # Chandra sometimes emits <img> tags in text blocks; they carry no image.
    for image in soup.find_all("img"):
        image.decompose()
    html = str(soup).strip()
    if label == "Text" and not re.search(r"<.+>", html):
        html = f"<p>{html}</p>"
    markdown = _tables_to_markdown(converter.convert(html), converter)
    # A bullet glyph or number kept inside a list item would follow the Markdown marker.
    markdown = _DOUBLED_BULLET.sub(r"\1", markdown)
    markdown = _DOUBLED_NUMBER.sub(r"\1\2. ", markdown)
    return markdown.strip()


def page_markdown(
    raw: str,
    page_image,
    page_number: int,
    mode: str,
    output_path: Path,
    assets_dir: Path,
    assets: list[Path],
    converter=None,
) -> str:
    """Turn one page of Chandra's layout HTML into Markdown.

    Page headers and footers are dropped, as with Docling. For pictures only
    the crop is kept: Chandra is prompted to describe images, chart values
    and diagrams in its own words, and that text is not in the source.
    """
    from chandra.output import parse_layout

    converter = converter or _markdown_converter()
    blocks = []
    pictures = 0
    for block in parse_layout(raw, page_image):
        if block.label in SKIPPED_LABELS:
            continue
        if block.label in PICTURE_LABELS:
            pictures += 1
            blocks.append(
                _picture_markdown(
                    page_image, block.bbox, page_number, pictures, mode, output_path, assets_dir, assets
                )
            )
            continue
        markdown = _text_block_markdown(block.label, block.content, converter)
        if markdown:
            blocks.append(markdown)
    return "\n\n".join(blocks)


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
                include_images=False,
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
            parts.append(
                page_markdown(result.raw, image, page_number, images, output_path, assets_dir, assets, converter)
            )
        del pages

    markdown = "\n\n".join(part for part in parts if part)
    return markdown.rstrip() + "\n", assets, warnings
