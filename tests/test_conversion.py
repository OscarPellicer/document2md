from __future__ import annotations

import subprocess
import sys
from pathlib import Path
from types import SimpleNamespace

from docx import Document
from openpyxl import Workbook
from PIL import Image
from pptx import Presentation
from pptx.util import Inches

from document2md import converter as converter_module
from document2md.converter import ConversionOptions, convert_file, convert_files


ROOT = Path(__file__).resolve().parents[1]
SKILL_SCRIPT = ROOT / "skills" / "document2md" / "scripts" / "convert_document.py"


def test_docx_preserves_paragraph_table_order(tmp_path: Path) -> None:
    document = Document()
    document.add_heading("Real heading", level=1)
    document.add_paragraph("Before table.")
    table = document.add_table(rows=2, cols=2)
    table.cell(0, 0).text = "A"
    table.cell(0, 1).text = "B"
    table.cell(1, 0).text = "1"
    table.cell(1, 1).text = "2"
    document.add_paragraph("After table.")
    source = tmp_path / "sample.docx"
    document.save(source)

    result = convert_file(source, options=ConversionOptions(force=True))
    markdown = result.output_path.read_text(encoding="utf-8")

    assert markdown.index("# Real heading") < markdown.index("Before table.")
    assert "| A | B |" in markdown
    assert markdown.index("| A | B |") < markdown.index("After table.")


def test_docx_does_not_infer_bold_heading_and_keeps_merged_table_grid(tmp_path: Path) -> None:
    document = Document()
    paragraph = document.add_paragraph()
    paragraph.add_run("Bold but not a heading").bold = True
    table = document.add_table(rows=2, cols=3)
    table.cell(0, 0).merge(table.cell(0, 1)).text = "Merged"
    table.cell(0, 2).text = "Third"
    table.cell(1, 0).text = "A"
    table.cell(1, 1).text = "B"
    table.cell(1, 2).text = "C"
    source = tmp_path / "merged.docx"
    document.save(source)

    markdown = convert_file(source, options=ConversionOptions(force=True)).output_path.read_text(encoding="utf-8")
    assert "**Bold but not a heading**" in markdown
    assert "# Bold but not a heading" not in markdown
    assert "| Merged |  | Third |" in markdown


def test_docx_images_are_only_saved_when_requested(tmp_path: Path) -> None:
    image_path = tmp_path / "figure.png"
    Image.new("RGB", (20, 20), "red").save(image_path)
    document = Document()
    document.add_paragraph("Text")
    document.add_picture(str(image_path))
    source = tmp_path / "pictured.docx"
    document.save(source)

    placeholder_output = tmp_path / "placeholder.md"
    placeholder = convert_file(
        source,
        placeholder_output,
        ConversionOptions(images="placeholder", force=True),
    )
    assert placeholder.assets == []
    assert not (tmp_path / "placeholder_assets").exists()

    extracted_output = tmp_path / "extracted.md"
    extracted = convert_file(
        source,
        extracted_output,
        ConversionOptions(images="extract", force=True),
    )
    assert len(extracted.assets) == 1
    assert extracted.assets[0].is_file()
    assert "extracted_assets/" in extracted_output.read_text(encoding="utf-8")


def test_pptx_converts_slide_text_and_table(tmp_path: Path) -> None:
    presentation = Presentation()
    slide = presentation.slides.add_slide(presentation.slide_layouts[5])
    slide.shapes.title.text = "Overview"
    box = slide.shapes.add_textbox(Inches(1), Inches(1.5), Inches(4), Inches(1))
    box.text = "Body text"
    table = slide.shapes.add_table(2, 2, Inches(1), Inches(3), Inches(5), Inches(1.5)).table
    table.cell(0, 0).text = "Metric"
    table.cell(0, 1).text = "Value"
    table.cell(1, 0).text = "Count"
    table.cell(1, 1).text = "3"
    source = tmp_path / "slides.pptx"
    presentation.save(source)

    markdown = convert_file(source, options=ConversionOptions(force=True)).output_path.read_text(encoding="utf-8")
    assert "# Slide 1: Overview" in markdown
    assert "Body text" in markdown
    assert "| Metric | Value |" in markdown


def test_xlsx_preserves_formulas_and_sheet_boundaries(tmp_path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Data"
    sheet.append(["Item", "Value"])
    sheet.append(["Total", "=1+2"])
    hidden = workbook.create_sheet("Hidden")
    hidden.sheet_state = "hidden"
    hidden["A1"] = "Secret"
    source = tmp_path / "book.xlsx"
    workbook.save(source)

    markdown = convert_file(source, options=ConversionOptions(force=True)).output_path.read_text(encoding="utf-8")
    assert "# Sheet: Data" in markdown
    assert "=1+2" in markdown
    assert "Hidden" not in markdown


def test_skill_script_is_runnable_without_package_install(tmp_path: Path) -> None:
    workbook = Workbook()
    workbook.active.append(["A", "B"])
    source = tmp_path / "book.xlsx"
    workbook.save(source)

    result = subprocess.run(
        [sys.executable, str(SKILL_SCRIPT), str(source), "--force"],
        check=True,
        capture_output=True,
        text=True,
    )
    assert str(tmp_path / "book.md") in result.stdout
    assert "| A | B |" in (tmp_path / "book.md").read_text(encoding="utf-8")


def test_multiple_pdfs_share_one_docling_converter(tmp_path: Path, monkeypatch) -> None:
    sources = [tmp_path / "one.pdf", tmp_path / "two.pdf"]
    for source in sources:
        source.write_bytes(b"%PDF-test")
    targets = [tmp_path / "one.md", tmp_path / "two.md"]
    created = []

    class FakeConverter:
        def convert_all(self, paths):
            assert list(paths) == sources
            return iter(
                [
                    SimpleNamespace(document=SimpleNamespace(name="one")),
                    SimpleNamespace(document=SimpleNamespace(name="two")),
                ]
            )

    def fake_create(options):
        created.append(options)
        return FakeConverter(), object()

    def fake_serialize(document, output_path, options, image_ref_mode):
        return f"# {document.name}\n", [], []

    monkeypatch.setattr(converter_module, "_create_pdf_converter", fake_create)
    monkeypatch.setattr(converter_module, "_serialize_pdf_document", fake_serialize)

    results = convert_files(sources, targets, ConversionOptions(force=True))

    assert len(created) == 1
    assert [result.output_path for result in results] == targets
    assert targets[0].read_text(encoding="utf-8") == "# one\n"
    assert targets[1].read_text(encoding="utf-8") == "# two\n"


def test_pdf_extract_mode_uses_accompanying_artifacts_folder(tmp_path: Path) -> None:
    class FakeDocument:
        def save_as_markdown(self, filename, *, image_mode, artifacts_dir):
            artifacts_dir.mkdir(parents=True, exist_ok=True)
            (artifacts_dir / "picture.png").write_bytes(b"png")
            filename.write_text("![picture](sample_assets/picture.png)\n", encoding="utf-8")

    output = tmp_path / "sample.md"
    markdown, assets, warnings = converter_module._serialize_pdf_document(
        FakeDocument(),
        output,
        ConversionOptions(images="extract"),
        SimpleNamespace(REFERENCED="referenced"),
    )

    assert warnings == []
    assert markdown.startswith("![picture]")
    assert assets == [tmp_path / "sample_assets" / "picture.png"]
