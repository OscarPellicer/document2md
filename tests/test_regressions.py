"""Regressions found while converting a real accreditation dossier.

Each test here pins behaviour that silently went wrong in the field: image
links that broke when a folder moved, a rewrite that was never written, a batch
that lost work because one file failed, and formats the tool used to reject.
"""

from __future__ import annotations

import subprocess
import sys
import zipfile
from pathlib import Path
from types import SimpleNamespace

import pytest
from openpyxl import Workbook

from document2md import converter as converter_module
from document2md.converter import ConversionOptions, convert_file, convert_files


ROOT = Path(__file__).resolve().parents[1]
SKILL_SCRIPT = ROOT / "skills" / "document2md" / "scripts" / "convert_document.py"


class FakeDoclingDocument:
    """Stands in for a docling document, reproducing its link behaviour.

    docling_core only emits relative image links when the artifacts directory it
    receives is relative; given an absolute one it writes an absolute,
    percent-encoded URI (docling_core `_get_output_paths`).
    """

    def __init__(self, picture_count: int = 1) -> None:
        self.picture_count = picture_count

    def save_as_markdown(self, filename: Path, *, image_mode, artifacts_dir: Path) -> None:
        if artifacts_dir.is_absolute():
            resolved = artifacts_dir
            link_for = lambda target: _percent_encode(target)
        else:
            resolved = filename.parent / artifacts_dir
            link_for = lambda target: _percent_encode(target.relative_to(filename.parent))
        resolved.mkdir(parents=True, exist_ok=True)
        lines = []
        for index in range(self.picture_count):
            target = resolved / f"image_{index:06d}.png"
            target.write_bytes(b"png")
            lines.append(f"![Image]({link_for(target)})")
        filename.write_text("\n\n".join(lines) + "\n", encoding="utf-8")

    def export_to_markdown(self, *, image_mode) -> str:
        return "text only"


def _percent_encode(path: Path) -> str:
    from urllib.parse import quote

    return quote(path.as_posix())


def _install_fake_docling(monkeypatch, page_count: int, picture_count: int = 1) -> None:
    class FakeConverter:
        def convert(self, stream, *, page_range):
            return SimpleNamespace(document=FakeDoclingDocument(picture_count))

    monkeypatch.setattr(
        converter_module,
        "_create_pdf_converter",
        lambda options: (FakeConverter(), SimpleNamespace(REFERENCED="referenced", PLACEHOLDER="placeholder", EMBEDDED="embedded")),
    )
    monkeypatch.setattr(converter_module, "_pdf_page_count", lambda path: page_count)


# --- Image links must survive the folder being moved ------------------------


@pytest.mark.parametrize("page_count", [1, 25], ids=["single-batch", "multi-batch"])
def test_extracted_pdf_image_links_are_relative(tmp_path: Path, monkeypatch, page_count: int) -> None:
    _install_fake_docling(monkeypatch, page_count=page_count)
    source = tmp_path / "guía de trámite.pdf"
    source.write_bytes(b"%PDF-test")
    output = tmp_path / "out" / "guide.md"

    convert_file(source, output, ConversionOptions(images="extract", force=True))
    markdown = output.read_text(encoding="utf-8")

    assert "](/" not in markdown, "image link is an absolute path"
    assert "file://" not in markdown
    assert str(tmp_path) not in markdown
    assert "guide_assets/" in markdown


def test_relativize_rewrites_absolute_links_and_leaves_others_alone(tmp_path: Path) -> None:
    output = tmp_path / "doc.md"
    assets = tmp_path / "doc_assets"
    absolute = f"![Image]({_percent_encode(assets / 'a b.png')})"
    file_uri = f"![Image](file://{_percent_encode(assets / 'c.png')})"
    already_relative = "![Image](doc_assets/d.png)"
    remote = "![Image](https://example.org/e.png)"
    embedded = "![Image](data:image/png;base64,AAAA)"

    rewritten = converter_module._relativize_markdown_links(
        "\n".join([absolute, file_uri, already_relative, remote, embedded]),
        output,
    )

    assert "![Image](doc_assets/a%20b.png)" in rewritten
    assert "![Image](doc_assets/c.png)" in rewritten
    assert already_relative in rewritten
    assert remote in rewritten
    assert embedded in rewritten


def test_relative_links_still_resolve_after_the_folder_moves(tmp_path: Path, monkeypatch) -> None:
    _install_fake_docling(monkeypatch, page_count=1)
    source = tmp_path / "doc.pdf"
    source.write_bytes(b"%PDF-test")
    original = tmp_path / "here" / "doc.md"
    convert_file(source, original, ConversionOptions(images="extract", force=True))

    moved = tmp_path / "somewhere else"
    (tmp_path / "here").rename(moved)
    markdown = (moved / "doc.md").read_text(encoding="utf-8")

    link = markdown.split("](", 1)[1].split(")", 1)[0]
    from urllib.parse import unquote

    assert (moved / unquote(link)).is_file()


# --- A conversion that reports success must be on disk ----------------------


@pytest.mark.parametrize("page_count", [1, 25], ids=["single-batch", "multi-batch"])
def test_force_overwrites_an_existing_markdown_in_extract_mode(
    tmp_path: Path, monkeypatch, page_count: int
) -> None:
    """Regression: the multi-batch extract path skipped the write and kept stale text."""
    _install_fake_docling(monkeypatch, page_count=page_count)
    source = tmp_path / "doc.pdf"
    source.write_bytes(b"%PDF-test")
    output = tmp_path / "doc.md"
    output.write_text("STALE CONTENT FROM A PREVIOUS RUN\n", encoding="utf-8")

    convert_file(source, output, ConversionOptions(images="extract", force=True))

    assert "STALE" not in output.read_text(encoding="utf-8")


def test_force_overwrites_in_extract_mode_through_the_batch_path(tmp_path: Path, monkeypatch) -> None:
    _install_fake_docling(monkeypatch, page_count=25)
    source = tmp_path / "doc.pdf"
    source.write_bytes(b"%PDF-test")
    output = tmp_path / "doc.md"
    output.write_text("STALE CONTENT FROM A PREVIOUS RUN\n", encoding="utf-8")

    convert_files([source], [output], ConversionOptions(images="extract", force=True))

    assert "STALE" not in output.read_text(encoding="utf-8")


# --- One bad input must not cost the rest of the batch ----------------------


def test_a_failing_input_does_not_discard_the_rest_of_the_batch(tmp_path: Path) -> None:
    good_first = tmp_path / "first.xlsx"
    good_last = tmp_path / "last.xlsx"
    for path in (good_first, good_last):
        workbook = Workbook()
        workbook.active.append(["A", "B"])
        workbook.save(path)
    broken = tmp_path / "broken.xlsx"
    broken.write_bytes(b"not a workbook at all")

    batch = convert_files(
        [good_first, broken, good_last],
        [tmp_path / "first.md", tmp_path / "broken.md", tmp_path / "last.md"],
        ConversionOptions(force=True),
    )

    assert [result.input_path.name for result in batch.results] == ["first.xlsx", "last.xlsx"]
    assert [failure.input_path.name for failure in batch.failures] == ["broken.xlsx"]
    assert (tmp_path / "last.md").is_file()


def test_strict_mode_raises_on_the_first_failure(tmp_path: Path) -> None:
    broken = tmp_path / "broken.xlsx"
    broken.write_bytes(b"not a workbook at all")

    with pytest.raises(Exception):
        convert_files([broken], [tmp_path / "broken.md"], ConversionOptions(force=True), strict=True)


def test_cli_reports_failures_and_exits_non_zero(tmp_path: Path) -> None:
    workbook = Workbook()
    workbook.active.append(["A", "B"])
    good = tmp_path / "good.xlsx"
    workbook.save(good)
    broken = tmp_path / "broken.xlsx"
    broken.write_bytes(b"not a workbook at all")

    result = subprocess.run(
        [sys.executable, str(SKILL_SCRIPT), str(good), str(broken), "--force", "--no-bootstrap"],
        capture_output=True,
        text=True,
    )

    assert result.returncode == 1
    assert str(tmp_path / "good.md") in result.stdout
    assert "broken.xlsx" in result.stderr
    assert (tmp_path / "good.md").is_file()


# --- OCR mode ---------------------------------------------------------------


def test_ocr_mode_is_validated() -> None:
    with pytest.raises(ValueError):
        convert_file(__file__, options=ConversionOptions(ocr="sometimes"))


@pytest.mark.parametrize(
    ("mode", "do_ocr", "full_page"),
    [("auto", True, False), ("off", False, False), ("force", True, True)],
)
def test_ocr_mode_reaches_the_docling_pipeline(monkeypatch, mode: str, do_ocr: bool, full_page: bool) -> None:
    docling = pytest.importorskip("docling.datamodel.pipeline_options")
    captured = {}

    class FakeDocumentConverter:
        def __init__(self, *, allowed_formats, format_options):
            captured["options"] = format_options

    monkeypatch.setattr("docling.document_converter.DocumentConverter", FakeDocumentConverter)
    converter_module._create_pdf_converter(ConversionOptions(ocr=mode))

    pipeline = next(iter(captured["options"].values())).pipeline_options
    assert isinstance(pipeline, docling.PdfPipelineOptions)
    assert pipeline.do_ocr is do_ocr
    assert _is_full_page(pipeline.ocr_options) is full_page


def _is_full_page(ocr_options) -> bool:
    mode = getattr(ocr_options, "mode", None)
    if mode is not None:
        return str(getattr(mode, "value", mode)) == "full_page"
    return bool(getattr(ocr_options, "force_full_page_ocr", False))
