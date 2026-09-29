"""Acrobat comment/annotation extraction for the PDF conversion path.

Docling (the library behind PDF conversion, see ``converter._convert_pdf``)
never reads a PDF's ``/Annots``, so review comments made in Adobe Acrobat
(highlight annotations with attached notes, sticky notes) are silently
dropped from the converted Markdown. This module extracts them independently
via pypdfium2 -- already a hard dependency of this project -- and inserts
them as ``[^cN]`` footnote-style markers at the point in the Markdown where
each comment is anchored, without needing any changes to Docling itself.
"""

from __future__ import annotations

import ctypes
import re
from itertools import count
from typing import Any, Iterator

import pypdfium2.raw as pdfium_raw

# Annotation subtypes that highlight/mark up an underlying text region.
_MARKUP_SUBTYPES = {
    pdfium_raw.FPDF_ANNOT_HIGHLIGHT,
    pdfium_raw.FPDF_ANNOT_UNDERLINE,
    pdfium_raw.FPDF_ANNOT_STRIKEOUT,
    pdfium_raw.FPDF_ANNOT_SQUIGGLY,
}
# Annotation subtypes anchored to a point rather than a text region (sticky notes).
_POINT_SUBTYPES = {pdfium_raw.FPDF_ANNOT_TEXT}

# Vertical padding (in PDF points) used to approximate "the nearest line" for
# point-anchored (sticky note) annotations, which have no underlying text region.
_POINT_ANCHOR_BAND = 8


def _get_annot_string(annot, key: str) -> str | None:
    """Read a UTF-16LE string value (e.g. /Contents, /T) off a pdfium annotation."""
    key_bytes = key.encode("utf-8")
    buflen = pdfium_raw.FPDFAnnot_GetStringValue(annot, key_bytes, None, 0)
    if buflen <= 2:  # empty string is just the null terminator
        return None
    buf = ctypes.create_string_buffer(buflen)
    pdfium_raw.FPDFAnnot_GetStringValue(
        annot, key_bytes, ctypes.cast(buf, ctypes.POINTER(ctypes.c_ushort)), buflen
    )
    text = bytes(buf.raw[:buflen]).decode("utf-16-le", errors="ignore").rstrip("\x00")
    return text if text.strip() else None


def _extract_page_comments(pdfium_page: Any) -> list[dict]:
    """Extract Acrobat comment annotations from a pypdfium2 page.

    Returns a list of ``{"contents": str, "author": str | None, "anchor_text": str | None}``
    dicts, one per comment-bearing annotation (Popups, Links, and markup
    annotations without an attached note are skipped).
    """
    raw_page = pdfium_page.raw
    count_annots = pdfium_raw.FPDFPage_GetAnnotCount(raw_page)
    comments: list[dict] = []
    textpage = None
    _, page_height = pdfium_page.get_size()

    for i in range(count_annots):
        annot = pdfium_raw.FPDFPage_GetAnnot(raw_page, i)
        try:
            contents = _get_annot_string(annot, "Contents")
            if not contents:
                continue

            subtype = pdfium_raw.FPDFAnnot_GetSubtype(annot)
            rect = pdfium_raw.FS_RECTF()
            if not pdfium_raw.FPDFAnnot_GetRect(annot, ctypes.byref(rect)):
                continue

            if textpage is None:
                textpage = pdfium_page.get_textpage()

            anchor_text: str | None = None
            try:
                if subtype in _MARKUP_SUBTYPES:
                    anchor_text = textpage.get_text_bounded(
                        left=rect.left, bottom=rect.bottom, right=rect.right, top=rect.top
                    )
                elif subtype in _POINT_SUBTYPES:
                    anchor_text = textpage.get_text_bounded(
                        left=0,
                        right=pdfium_page.get_size()[0],
                        bottom=max(0, rect.top - _POINT_ANCHOR_BAND),
                        top=min(page_height, rect.top + _POINT_ANCHOR_BAND),
                    )
                else:
                    continue  # unsupported/unplaceable subtype
            except Exception:
                anchor_text = None

            comments.append(
                {
                    "contents": contents.strip(),
                    "author": _get_annot_string(annot, "T"),
                    "anchor_text": anchor_text.strip() if anchor_text else None,
                }
            )
        finally:
            pdfium_raw.FPDFPage_CloseAnnot(annot)

    return comments


def _insert_comment_markers(
    text: str, comments: list[dict], marker_counter: Iterator[int]
) -> tuple[str, list[tuple[int, dict]]]:
    """Insert ``[^cN]`` markers into ``text`` at each comment's anchor location.

    Falls back to appending at the end of ``text`` when the anchor can't be
    matched (e.g. Docling reformatted the surrounding text), so a comment is
    never silently lost -- only imprecisely placed.

    Returns ``(new_text, [(marker_number, comment), ...])``.
    """
    footnotes: list[tuple[int, dict]] = []

    for comment in comments:
        n = next(marker_counter)
        marker = f"[^c{n}]"
        anchor_text = comment.get("anchor_text")
        inserted = False

        if anchor_text:
            pattern = r"\s+".join(re.escape(part) for part in anchor_text.split())
            if pattern:
                match = re.search(pattern, text)
                if match:
                    end = match.end()
                    # Snap forward past any remaining non-whitespace characters,
                    # so a highlight that stops mid-word (e.g. Acrobat selection
                    # ending on "p" of "problem") doesn't split the word with
                    # the marker -- the whole word gets the marker instead.
                    while end < len(text) and not text[end].isspace():
                        end += 1
                    text = text[:end] + marker + text[end:]
                    inserted = True

        if not inserted:
            text = text.rstrip() + " " + marker

        footnotes.append((n, comment))

    return text, footnotes


def render_comments_section(footnotes: list[tuple[int, int, dict]]) -> str:
    """Render the trailing '## Comments' section for a list of
    ``(marker_number, page_number, comment)`` tuples, sorted by marker number."""
    lines = ["## Comments", ""]
    for marker_number, page_number, comment in sorted(footnotes, key=lambda item: item[0]):
        author = comment.get("author")
        prefix = f"{author}: " if author else ""
        lines.append(f"[^c{marker_number}]: (p.{page_number}) {prefix}{comment['contents']}")
    return "\n\n---\n\n" + "\n".join(lines)


def new_marker_counter() -> Iterator[int]:
    return count(1)
