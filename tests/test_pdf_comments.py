from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace

import pypdfium2 as pdfium

from document2md import converter as converter_module
from document2md import pdf_comments
from document2md.converter import ConversionOptions, convert_file

FIXTURES = Path(__file__).resolve().parent / "fixtures"
ANNOTATED_PDF = FIXTURES / "annotated_sample.pdf"
COMMENT_THREAD_PDF = FIXTURES / "comment_thread.pdf"


def test_extract_page_comments_reads_highlight_contents_author_and_anchor() -> None:
    doc = pdfium.PdfDocument(ANNOTATED_PDF)
    try:
        comments = pdf_comments._extract_page_comments(doc[0])
    finally:
        doc.close()

    assert len(comments) == 1
    comment = comments[0]
    assert comment["contents"] == "is this seasonal pattern statistically significant?"
    assert comment["author"] == "Test Reviewer"
    assert comment["anchor_text"] == "precipitation dataset shows a clear seasonal pattern"


def test_insert_comment_markers_lands_marker_at_anchor_text() -> None:
    text = (
        "This document exists only to exercise the PDF comment extraction feature.\n"
        "The precipitation dataset shows a clear seasonal pattern across all stations."
    )
    comments = [
        {
            "contents": "is this seasonal pattern statistically significant?",
            "author": "Test Reviewer",
            "anchor_text": "precipitation dataset shows a clear seasonal pattern",
        }
    ]

    new_text, footnotes = pdf_comments._insert_comment_markers(
        text, comments, pdf_comments.new_marker_counter()
    )

    assert "precipitation dataset shows a clear seasonal pattern[^c1]" in new_text
    assert footnotes == [(1, comments[0])]


def test_insert_comment_markers_snaps_to_word_boundary_when_anchor_ends_mid_word() -> None:
    """Regression: an Acrobat highlight that stops mid-word (e.g. selection ending
    on "p" of "problem") must not split the word with the marker."""
    text = "Unmeasured confounding is a crucial problem in observational studies."
    comments = [{"contents": "test1", "author": "natemankovich", "anchor_text": "crucial p"}]

    new_text, _ = pdf_comments._insert_comment_markers(
        text, comments, pdf_comments.new_marker_counter()
    )

    assert "crucial problem[^c1]" in new_text
    assert "p[^c1]roblem" not in new_text


def test_extract_page_comments_reads_a_thread_of_adjacent_highlights_in_order() -> None:
    """comment_thread.pdf has one highlight from the base fixture plus three more,
    consecutive, back-to-back highlights on the same sentence (a comment thread)."""
    doc = pdfium.PdfDocument(COMMENT_THREAD_PDF)
    try:
        comments = pdf_comments._extract_page_comments(doc[0])
    finally:
        doc.close()

    assert len(comments) == 4
    # Extraction order follows annotation order in the PDF, i.e. reading order.
    assert [c["anchor_text"] for c in comments] == [
        "precipitation dataset shows a clear seasonal pattern",
        "Reviewers often",
        "leave highlight comments",
        "on specific phrases",
    ]
    assert [c["author"] for c in comments[1:]] == ["Alice", "Bob", "Alice"]


def test_insert_comment_markers_places_adjacent_thread_markers_at_each_own_phrase() -> None:
    """Three consecutive, non-overlapping highlights in the same sentence must each
    get their own marker at their own phrase, not bleed into a neighbor's."""
    text = "Reviewers often leave highlight comments on specific phrases like this one."
    comments = [
        {"contents": "reply 1", "author": "Alice", "anchor_text": "Reviewers often"},
        {"contents": "reply 2", "author": "Bob", "anchor_text": "leave highlight comments"},
        {"contents": "reply 3", "author": "Alice", "anchor_text": "on specific phrases"},
    ]

    new_text, footnotes = pdf_comments._insert_comment_markers(
        text, comments, pdf_comments.new_marker_counter()
    )

    assert new_text == (
        "Reviewers often[^c1] leave highlight comments[^c2] on specific phrases[^c3]"
        " like this one."
    )
    assert [n for n, _ in footnotes] == [1, 2, 3]


def test_insert_comment_markers_stacks_a_reply_thread_on_the_same_anchor_in_order() -> None:
    """A literal Acrobat reply thread: multiple comments all anchored to the exact
    same highlighted phrase. Markers must stack in comment order, not reverse or
    collide, at the single shared anchor point."""
    text = "Reviewers often leave highlight comments on specific phrases like this one."
    comments = [
        {"contents": "reply 1", "author": "Alice", "anchor_text": "leave highlight comments"},
        {"contents": "reply 2", "author": "Bob", "anchor_text": "leave highlight comments"},
        {"contents": "reply 3", "author": "Alice", "anchor_text": "leave highlight comments"},
    ]

    new_text, footnotes = pdf_comments._insert_comment_markers(
        text, comments, pdf_comments.new_marker_counter()
    )

    assert "leave highlight comments[^c1][^c2][^c3] on specific phrases" in new_text
    assert [n for n, _ in footnotes] == [1, 2, 3]


def test_pdf_conversion_handles_comment_thread_end_to_end(
    tmp_path: Path, monkeypatch
) -> None:
    """Full pipeline (annotation extraction -> marker insertion -> Comments section)
    against a real fixture with four comments, three of them a back-to-back thread."""
    output = tmp_path / "result.md"

    class FakeDocument:
        def export_to_markdown(self, *, image_mode):
            return (
                "Sample Report\n\n"
                "This document exists only to exercise the PDF comment extraction feature.\n\n"
                "The precipitation dataset shows a clear seasonal pattern across all stations.\n\n"
                "Reviewers often leave highlight comments on specific phrases like this one.\n\n"
                "A second paragraph follows to give the layout more than one block of text."
            )

    class FakeConverter:
        def convert(self, stream, *, page_range):
            return SimpleNamespace(document=FakeDocument())

    monkeypatch.setattr(
        converter_module,
        "_create_pdf_converter",
        lambda options: (
            FakeConverter(),
            SimpleNamespace(PLACEHOLDER="placeholder", EMBEDDED="embedded"),
        ),
    )
    monkeypatch.setattr(
        converter_module, "_pdf_stream", lambda path, batch_number=0: SimpleNamespace()
    )

    convert_file(COMMENT_THREAD_PDF, output, ConversionOptions(force=True))
    markdown = output.read_text(encoding="utf-8")

    assert (
        "Reviewers often[^c2] leave highlight comments[^c3] on specific phrases[^c4]"
        " like this one."
        in markdown
    )
    assert "[^c2]: (p.1) Alice: First reply: agreed, flagging for follow-up." in markdown
    assert "[^c3]: (p.1) Bob: Second reply: same here, +1." in markdown
    assert "[^c4]: (p.1) Alice: Third reply: let's resolve this thread." in markdown


def test_insert_comment_markers_falls_back_to_end_when_anchor_not_found() -> None:
    text = "Nothing here matches the anchor text at all."
    comments = [{"contents": "a stray comment", "author": None, "anchor_text": "missing phrase"}]

    new_text, footnotes = pdf_comments._insert_comment_markers(
        text, comments, pdf_comments.new_marker_counter()
    )

    # The comment is never silently dropped, even when its anchor can't be matched.
    assert new_text.rstrip().endswith("[^c1]")
    assert footnotes == [(1, comments[0])]


def test_render_comments_section_formats_marker_page_author_and_text() -> None:
    footnotes = [
        (2, 4, {"contents": "second", "author": None}),
        (1, 1, {"contents": "first", "author": "Reviewer"}),
    ]

    section = pdf_comments.render_comments_section(footnotes)

    assert section.startswith("\n\n---\n\n## Comments\n\n")
    # Sorted by marker number, regardless of input order.
    assert section.index("[^c1]:") < section.index("[^c2]:")
    assert "[^c1]: (p.1) Reviewer: first" in section
    assert "[^c2]: (p.4) second" in section


def test_pdf_conversion_inserts_comment_markers_without_real_docling(
    tmp_path: Path, monkeypatch
) -> None:
    """The wiring in _convert_pdf_with_converter works even without Docling installed:
    only the annotation extraction (real pypdfium2, real fixture) needs to be genuine;
    Docling's own conversion is faked, exactly like the project's existing PDF tests."""
    output = tmp_path / "result.md"

    class FakeDocument:
        def export_to_markdown(self, *, image_mode):
            return (
                "This document exists only to exercise the PDF comment extraction feature.\n\n"
                "The precipitation dataset shows a clear seasonal pattern across all stations."
            )

    class FakeConverter:
        def convert(self, stream, *, page_range):
            return SimpleNamespace(document=FakeDocument())

    monkeypatch.setattr(
        converter_module,
        "_create_pdf_converter",
        lambda options: (
            FakeConverter(),
            SimpleNamespace(PLACEHOLDER="placeholder", EMBEDDED="embedded"),
        ),
    )
    # Avoid importing docling_core (not installed in this test environment);
    # the fake converter above never inspects the stream it's handed anyway.
    monkeypatch.setattr(
        converter_module, "_pdf_stream", lambda path, batch_number=0: SimpleNamespace()
    )

    convert_file(ANNOTATED_PDF, output, ConversionOptions(force=True))
    markdown = output.read_text(encoding="utf-8")

    assert "precipitation dataset shows a clear seasonal pattern[^c1]" in markdown
    assert "## Comments" in markdown
    assert "[^c1]: (p.1) Test Reviewer: is this seasonal pattern statistically significant?" in markdown


def test_pdf_conversion_without_comments_is_unaffected(tmp_path: Path, monkeypatch) -> None:
    """A PDF with no annotations produces byte-identical output to before this change."""
    source = tmp_path / "plain.pdf"
    # Build a real, valid, unannotated single-page PDF via pypdfium2 alone.
    doc = pdfium.PdfDocument.new()
    doc.new_page(200, 200)
    doc.save(source)
    output = tmp_path / "plain.md"

    class FakeDocument:
        def export_to_markdown(self, *, image_mode):
            return "Just some plain markdown."

    class FakeConverter:
        def convert(self, stream, *, page_range):
            return SimpleNamespace(document=FakeDocument())

    monkeypatch.setattr(
        converter_module,
        "_create_pdf_converter",
        lambda options: (
            FakeConverter(),
            SimpleNamespace(PLACEHOLDER="placeholder", EMBEDDED="embedded"),
        ),
    )
    monkeypatch.setattr(
        converter_module, "_pdf_stream", lambda path, batch_number=0: SimpleNamespace()
    )

    convert_file(source, output, ConversionOptions(force=True))
    markdown = output.read_text(encoding="utf-8")

    assert markdown == "Just some plain markdown.\n"
    assert "## Comments" not in markdown
