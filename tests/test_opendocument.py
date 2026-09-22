"""OpenDocument (.odt, .ods, .odp) extraction.

The fixtures are written as raw OpenDocument XML so the tests state exactly
which encoded structure is expected to survive, with no authoring library in
between.
"""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest

from document2md.converter import ConversionOptions, convert_file


NAMESPACES = (
    'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" '
    'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" '
    'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" '
    'xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" '
    'xmlns:xlink="http://www.w3.org/1999/xlink"'
)


def _write_opendocument(path: Path, body: str, media: dict[str, bytes] | None = None) -> Path:
    content = f'<?xml version="1.0" encoding="UTF-8"?>\n<office:document-content {NAMESPACES}><office:body>{body}</office:body></office:document-content>'
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("content.xml", content)
        for name, blob in (media or {}).items():
            archive.writestr(name, blob)
    return path


def test_odt_keeps_headings_paragraphs_tables_and_whitespace(tmp_path: Path) -> None:
    body = (
        "<office:text>"
        '<text:h text:outline-level="1">Autoinforme</text:h>'
        '<text:h text:outline-level="3">Apartado 2</text:h>'
        "<text:p>Antes de la tabla.</text:p>"
        '<text:p>Dos<text:s text:c="2"/>espacios y<text:tab/>tabulador.</text:p>'
        '<text:list><text:list-item><text:p>Primero</text:p></text:list-item>'
        "<text:list-item><text:p>Segundo</text:p></text:list-item></text:list>"
        "<table:table><table:table-row>"
        "<table:table-cell><text:p>Concepto</text:p></table:table-cell>"
        "<table:table-cell><text:p>Valor</text:p></table:table-cell>"
        "</table:table-row><table:table-row>"
        "<table:table-cell><text:p>Docencia</text:p></table:table-cell>"
        "<table:table-cell><text:p>120 h</text:p></table:table-cell>"
        "</table:table-row></table:table>"
        "<text:p>Despues de la tabla.</text:p>"
        "</office:text>"
    )
    source = _write_opendocument(tmp_path / "autoinforme.odt", body)

    markdown = convert_file(source, options=ConversionOptions(force=True)).output_path.read_text(encoding="utf-8")

    assert "# Autoinforme" in markdown
    assert "### Apartado 2" in markdown
    assert "- Primero" in markdown and "- Segundo" in markdown
    assert "| Concepto | Valor |" in markdown
    assert "| Docencia | 120 h |" in markdown
    assert markdown.index("Antes de la tabla.") < markdown.index("| Concepto | Valor |")
    assert markdown.index("| Concepto | Valor |") < markdown.index("Despues de la tabla.")
    assert "tabulador" in markdown


def test_odt_covered_cells_stay_empty_and_are_not_inferred(tmp_path: Path) -> None:
    body = (
        "<office:text><table:table><table:table-row>"
        '<table:table-cell table:number-columns-spanned="2"><text:p>Fusionada</text:p></table:table-cell>'
        "<table:covered-table-cell/>"
        "<table:table-cell><text:p>Tercera</text:p></table:table-cell>"
        "</table:table-row></table:table></office:text>"
    )
    source = _write_opendocument(tmp_path / "merged.odt", body)

    markdown = convert_file(source, options=ConversionOptions(force=True)).output_path.read_text(encoding="utf-8")

    assert "| Fusionada |  | Tercera |" in markdown


def test_odt_images_are_only_written_when_requested(tmp_path: Path) -> None:
    body = (
        "<office:text><text:p>Figura"
        '<draw:frame><draw:image xlink:href="Pictures/figura.png"/></draw:frame>'
        "</text:p></office:text>"
    )
    source = _write_opendocument(tmp_path / "figura.odt", body, {"Pictures/figura.png": b"png-bytes"})

    placeholder = convert_file(
        source, tmp_path / "placeholder.md", options=ConversionOptions(force=True)
    )
    assert placeholder.assets == []
    assert not (tmp_path / "placeholder_assets").exists()

    extracted = convert_file(
        source, tmp_path / "extracted.md", options=ConversionOptions(images="extract", force=True)
    )
    markdown = extracted.output_path.read_text(encoding="utf-8")
    assert len(extracted.assets) == 1
    assert extracted.assets[0].read_bytes() == b"png-bytes"
    link = markdown.split("](", 1)[1].split(")", 1)[0]
    assert not link.startswith("/")
    assert (tmp_path / link).is_file()


def test_ods_sheets_become_sections_and_repeats_expand(tmp_path: Path) -> None:
    body = (
        "<office:spreadsheet>"
        '<table:table table:name="Movilidad"><table:table-row>'
        "<table:table-cell><text:p>Centro</text:p></table:table-cell>"
        "<table:table-cell><text:p>Meses</text:p></table:table-cell>"
        "</table:table-row><table:table-row>"
        '<table:table-cell table:number-columns-repeated="2"><text:p>x</text:p></table:table-cell>'
        "</table:table-row></table:table>"
        '<table:table table:name="Oculta" table:display="false"><table:table-row>'
        "<table:table-cell><text:p>secreto</text:p></table:table-cell>"
        "</table:table-row></table:table>"
        "</office:spreadsheet>"
    )
    source = _write_opendocument(tmp_path / "datos.ods", body)

    result = convert_file(source, options=ConversionOptions(force=True))
    markdown = result.output_path.read_text(encoding="utf-8")

    assert "# Sheet: Movilidad" in markdown
    assert "| Centro | Meses |" in markdown
    assert "| x | x |" in markdown
    assert "secreto" not in markdown
    assert any("Oculta" in warning for warning in result.warnings)

    included = convert_file(
        source, tmp_path / "con_ocultas.md", options=ConversionOptions(force=True, include_hidden=True)
    )
    assert "secreto" in included.output_path.read_text(encoding="utf-8")


def test_odp_slides_become_sections_in_stored_order(tmp_path: Path) -> None:
    body = (
        "<office:presentation>"
        '<draw:page draw:name="Introduccion"><draw:frame><draw:text-box>'
        "<text:p>Primera linea</text:p><text:p>Segunda linea</text:p>"
        "</draw:text-box></draw:frame></draw:page>"
        '<draw:page draw:name="Resultados"><draw:frame><draw:text-box>'
        "<text:p>Contenido</text:p></draw:text-box></draw:frame></draw:page>"
        "</office:presentation>"
    )
    source = _write_opendocument(tmp_path / "charla.odp", body)

    markdown = convert_file(source, options=ConversionOptions(force=True)).output_path.read_text(encoding="utf-8")

    assert "# Slide 1: Introduccion" in markdown
    assert "# Slide 2: Resultados" in markdown
    assert markdown.index("Primera linea") < markdown.index("Segunda linea") < markdown.index("# Slide 2")


def test_unreadable_opendocument_file_is_rejected_clearly(tmp_path: Path) -> None:
    source = tmp_path / "roto.odt"
    source.write_bytes(b"not a zip archive")

    with pytest.raises(Exception):
        convert_file(source, options=ConversionOptions(force=True))


def test_embed_mode_stays_pdf_only(tmp_path: Path) -> None:
    source = _write_opendocument(tmp_path / "doc.odt", "<office:text><text:p>Hola</text:p></office:text>")

    with pytest.raises(ValueError):
        convert_file(source, options=ConversionOptions(images="embed", force=True))
