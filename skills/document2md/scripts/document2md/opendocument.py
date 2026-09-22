"""OpenDocument (.odt, .ods, .odp) reading without external dependencies.

OpenDocument files are zip archives holding a flat XML body. Reading them
directly keeps the skill's extraction contract: encoded structure is preserved
and nothing is inferred from appearance.
"""

from __future__ import annotations

import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree

TEXT_NS = "urn:oasis:names:tc:opendocument:xmlns:text:1.0"
TABLE_NS = "urn:oasis:names:tc:opendocument:xmlns:table:1.0"
OFFICE_NS = "urn:oasis:names:tc:opendocument:xmlns:office:1.0"
DRAW_NS = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
XLINK_NS = "http://www.w3.org/1999/xlink"

# A repeat count this large means "to the end of the sheet" rather than real data.
MAX_REPEAT = 1024


def _tag(element) -> str:
    return element.tag.rsplit("}", 1)[-1]


def _namespace(element) -> str:
    if element.tag.startswith("{"):
        return element.tag[1:].split("}", 1)[0]
    return ""


def _attribute(element, namespace: str, name: str, default=None):
    return element.get(f"{{{namespace}}}{name}", default)


def _text_content(element) -> str:
    """Flatten a text element, honouring OpenDocument whitespace encoding."""
    pieces: list[str] = []

    def walk(node) -> None:
        for child in node:
            name = _tag(child)
            if _namespace(child) == TEXT_NS and name == "s":
                count = int(_attribute(child, TEXT_NS, "c", "1") or 1)
                pieces.append(" " * count)
            elif _namespace(child) == TEXT_NS and name == "tab":
                pieces.append("\t")
            elif _namespace(child) == TEXT_NS and name == "line-break":
                pieces.append("\n")
            else:
                if child.text:
                    pieces.append(child.text)
                walk(child)
            if child.tail:
                pieces.append(child.tail)

    if element.text:
        pieces.append(element.text)
    walk(element)
    return "".join(pieces)


def _clean(text: str) -> str:
    return re.sub(r"[ \t\r\f\v]+", " ", text).strip()


def _list_items(element, depth: int = 0) -> list[str]:
    lines: list[str] = []
    for item in element:
        if _tag(item) != "list-item":
            continue
        for child in item:
            name = _tag(child)
            if name == "list":
                lines.extend(_list_items(child, depth + 1))
            else:
                text = _clean(_text_content(child))
                if text:
                    lines.append(f"{'  ' * depth}- {text}")
    return lines


def _table_rows(table) -> list[list[str]]:
    rows: list[list[str]] = []
    for row in table.iter(f"{{{TABLE_NS}}}table-row"):
        row_repeat = min(int(_attribute(row, TABLE_NS, "number-rows-repeated", "1") or 1), MAX_REPEAT)
        values: list[str] = []
        for cell in row:
            if _tag(cell) not in {"table-cell", "covered-table-cell"}:
                continue
            repeat = min(int(_attribute(cell, TABLE_NS, "number-columns-repeated", "1") or 1), MAX_REPEAT)
            if _tag(cell) == "covered-table-cell":
                # A cell covered by a merge holds no content of its own.
                values.extend([""] * repeat)
                continue
            paragraphs = [_clean(_text_content(child)) for child in cell if _tag(child) in {"p", "h"}]
            text = "<br>".join(part for part in paragraphs if part)
            values.extend([text] * repeat)
        while values and values[-1] == "":
            values.pop()
        rows.extend([list(values) for _ in range(row_repeat)])
    while rows and not any(cell for cell in rows[-1]):
        rows.pop()
    return rows


def _body(path: Path):
    with zipfile.ZipFile(path) as archive:
        try:
            content = archive.read("content.xml")
        except KeyError as error:
            raise ValueError(f"Not a readable OpenDocument file: {path}") from error
    root = ElementTree.fromstring(content)
    body = root.find(f"{{{OFFICE_NS}}}body")
    if body is None:
        raise ValueError(f"OpenDocument file has no body: {path}")
    return body


def _extract_images(path: Path, elements, assets_dir: Path, output_path: Path) -> tuple[dict[str, str], list[Path]]:
    """Copy referenced Pictures/* members beside the Markdown, keyed by archive path."""
    from .converter import _asset_target, _relative_link

    links: dict[str, str] = {}
    assets: list[Path] = []
    used: set[str] = set()
    hrefs = []
    for element in elements:
        href = _attribute(element, XLINK_NS, "href")
        if href and href not in hrefs:
            hrefs.append(href)
    if not hrefs:
        return links, assets
    with zipfile.ZipFile(path) as archive:
        names = set(archive.namelist())
        for href in hrefs:
            member = href.lstrip("./")
            if member not in names:
                continue
            assets_dir.mkdir(parents=True, exist_ok=True)
            source = Path(member)
            target = _asset_target(assets_dir, source.stem, source.suffix or ".bin", used)
            target.write_bytes(archive.read(member))
            assets.append(target)
            links[href] = _relative_link(target, output_path)
    return links, assets


def convert_text(path: Path, output_path: Path, extract_images: bool, assets_dir: Path) -> tuple[str, list[Path], list[str]]:
    """Convert an .odt document body to Markdown in source order."""
    from .converter import _render_table

    body = _body(path)
    text_body = body.find(f"{{{OFFICE_NS}}}text")
    if text_body is None:
        raise ValueError(f"Not an OpenDocument text document: {path}")

    image_elements = list(text_body.iter(f"{{{DRAW_NS}}}image"))
    links: dict[str, str] = {}
    assets: list[Path] = []
    if extract_images:
        links, assets = _extract_images(path, image_elements, assets_dir, output_path)

    blocks: list[str] = []

    def emit(element) -> None:
        name = _tag(element)
        if name == "h":
            text = _clean(_text_content(element))
            if text:
                level = max(1, min(6, int(_attribute(element, TEXT_NS, "outline-level", "1") or 1)))
                blocks.append(f"{'#' * level} {text}")
        elif name == "p":
            text = _clean(_text_content(element))
            if text:
                blocks.append(text)
            for image in element.iter(f"{{{DRAW_NS}}}image"):
                href = _attribute(image, XLINK_NS, "href")
                if href and href in links:
                    blocks.append(f"![{Path(links[href]).stem}]({links[href]})")
        elif name == "list":
            lines = _list_items(element)
            if lines:
                blocks.append("\n".join(lines))
        elif name == "table":
            rendered = _render_table(_table_rows(element))
            if rendered:
                blocks.append(rendered)
        elif name in {"section", "text-box", "frame"}:
            for child in element:
                emit(child)

    for element in text_body:
        emit(element)

    return "\n\n".join(blocks).rstrip() + "\n", assets, []


def convert_spreadsheet(path: Path, include_hidden: bool) -> tuple[str, list[Path], list[str]]:
    """Convert an .ods workbook, one section per sheet."""
    from .converter import _render_table

    body = _body(path)
    sheet_body = body.find(f"{{{OFFICE_NS}}}spreadsheet")
    if sheet_body is None:
        raise ValueError(f"Not an OpenDocument spreadsheet: {path}")

    output: list[str] = []
    warnings: list[str] = []
    for table in sheet_body.findall(f"{{{TABLE_NS}}}table"):
        name = _attribute(table, TABLE_NS, "name", "Sheet")
        style = table.get(f"{{{TABLE_NS}}}style-name", "")
        visible = _attribute(table, TABLE_NS, "display", "true") != "false"
        if not visible and not include_hidden:
            warnings.append(f"Skipped hidden sheet {name!r} (style {style!r}).")
            continue
        rows = _table_rows(table)
        if not rows:
            continue
        output.extend([f"# Sheet: {name}", "", _render_table(rows), ""])
    return "\n".join(output).rstrip() + "\n", [], warnings


def convert_presentation(path: Path, output_path: Path, extract_images: bool, assets_dir: Path) -> tuple[str, list[Path], list[str]]:
    """Convert an .odp deck, one section per slide, in stored order."""
    from .converter import _render_table

    body = _body(path)
    presentation = body.find(f"{{{OFFICE_NS}}}presentation")
    if presentation is None:
        raise ValueError(f"Not an OpenDocument presentation: {path}")

    pages = presentation.findall(f"{{{DRAW_NS}}}page")
    links: dict[str, str] = {}
    assets: list[Path] = []
    if extract_images:
        images = [image for page in pages for image in page.iter(f"{{{DRAW_NS}}}image")]
        links, assets = _extract_images(path, images, assets_dir, output_path)

    output: list[str] = []
    for number, page in enumerate(pages, start=1):
        title = _attribute(page, DRAW_NS, "name", "") or ""
        output.extend([f"# Slide {number}" + (f": {title}" if title else ""), ""])
        for frame in page:
            if _tag(frame) != "frame":
                continue
            for child in frame:
                name = _tag(child)
                if name == "text-box":
                    lines = []
                    for paragraph in child:
                        if _tag(paragraph) == "list":
                            lines.extend(_list_items(paragraph))
                        else:
                            text = _clean(_text_content(paragraph))
                            if text:
                                lines.append(text)
                    if lines:
                        output.extend(["\n".join(lines), ""])
                elif name == "table":
                    rendered = _render_table(_table_rows(child))
                    if rendered:
                        output.extend([rendered, ""])
                elif name == "image":
                    href = _attribute(child, XLINK_NS, "href")
                    if href and href in links:
                        output.extend([f"![{Path(links[href]).stem}]({links[href]})", ""])
    return "\n".join(output).rstrip() + "\n", assets, []
