"""The optional Chandra PDF backend.

The model itself is never loaded here: a fake engine returns Chandra-shaped
results so the Markdown post-processing, image modes and failure handling can
be checked without the 10 GB checkpoint.
"""

from __future__ import annotations

import importlib.util
from pathlib import Path
from types import SimpleNamespace
from urllib.parse import unquote

import pytest

from document2md import converter as converter_module
from document2md.converter import ConversionOptions, convert_files

pytest.importorskip("chandra")
pytest.importorskip("pypdfium2")
from PIL import Image  # noqa: E402

from document2md import chandra_backend  # noqa: E402


ROOT = Path(__file__).resolve().parents[1]
SKILL_SCRIPT = ROOT / "skills" / "document2md" / "scripts" / "convert_document.py"


def _blank_pdf(path: Path, pages: int = 2) -> Path:
    images = [Image.new("RGB", (200, 280), "white") for _ in range(pages)]
    images[0].save(path, save_all=True, append_images=images[1:])
    return path


class FakeManager:
    """Return one canned Chandra result per page, with one figure on page 1."""

    def __init__(self, fail_pages: set[int] = frozenset()):
        self.calls = 0
        self.fail_pages = fail_pages

    def generate(self, batch, **kwargs):
        self.calls += 1
        page = self.calls
        if page in self.fail_pages:
            return [SimpleNamespace(raw="", markdown="", images={}, error=True)]
        figure = Image.new("RGB", (10, 10), "red")
        markdown = (
            f"# Page {page}\n\nText with a citation<sup>1,2</sup> and $x^2$.\n\n"
            "![Bar chart of survival](abc_1_img.webp)\n\n"
            "<table><tr><th>A</th><th>B</th></tr><tr><td colspan=\"2\">wide</td></tr></table>"
        )
        return [SimpleNamespace(raw="<div>ok</div>", markdown=markdown, images={"abc_1_img.webp": figure}, error=False)]


def _engine(manager=None):
    return chandra_backend.ChandraEngine(manager or FakeManager(), "hf")


def _run(tmp_path: Path, monkeypatch, images: str, manager=None, pages: int = 2):
    source = _blank_pdf(tmp_path / "scan ñ.pdf", pages)
    engine = _engine(manager)
    monkeypatch.setattr(converter_module, "_create_pdf_engine", lambda options: (engine, None))
    options = ConversionOptions(pdf_backend="chandra", images=images, pdf_pages_per_batch=1)
    batch = convert_files([source], [tmp_path / "out" / "scan.md"], options)
    return batch


def test_html_tables_become_markdown_grids_with_spans_written_once() -> None:
    html = (
        "<table><tr><th>Gene</th><th colspan='2'>Result</th></tr>"
        "<tr><td rowspan='2'>KRAS</td><td>G12D</td><td>pos</td></tr>"
        "<tr><td>G12V</td><td>neg | weak</td></tr></table>"
    )
    markdown = chandra_backend._html_table_to_markdown(html, chandra_backend._markdown_converter())
    assert markdown.splitlines() == [
        "| Gene | Result |  |",
        "| --- | --- | --- |",
        "| KRAS | G12D | pos |",
        "|  | G12V | neg \\| weak |",
    ]


def test_extracted_figures_are_written_and_linked_relatively(tmp_path: Path, monkeypatch) -> None:
    batch = _run(tmp_path, monkeypatch, "extract")
    assert not batch.failures
    result = batch.results[0]
    text = result.output_path.read_text(encoding="utf-8")
    assert "<sup>1,2</sup>" in text and "$x^2$" in text
    assert "| A | B |" in text and "<table" not in text
    assert [path.name for path in result.assets] == ["page-0001-image-1.png", "page-0002-image-1.png"]
    for line in text.splitlines():
        if line.startswith("!["):
            link = unquote(line[line.index("(") + 1 : -1])
            assert not Path(link).is_absolute()
            assert (result.output_path.parent / link).is_file()


def test_placeholder_mode_writes_no_assets(tmp_path: Path, monkeypatch) -> None:
    result = _run(tmp_path, monkeypatch, "placeholder").results[0]
    text = result.output_path.read_text(encoding="utf-8")
    assert result.assets == []
    assert text.count("<!-- image -->") == 2
    assert "abc_1_img" not in text


def test_embed_mode_inlines_png_data(tmp_path: Path, monkeypatch) -> None:
    result = _run(tmp_path, monkeypatch, "embed").results[0]
    assert "](data:image/png;base64," in result.output_path.read_text(encoding="utf-8")


def test_a_page_chandra_cannot_read_is_marked_and_reported(tmp_path: Path, monkeypatch) -> None:
    result = _run(tmp_path, monkeypatch, "placeholder", FakeManager(fail_pages={2}), pages=3).results[0]
    text = result.output_path.read_text(encoding="utf-8")
    assert "<!-- page 2: not transcribed -->" in text
    assert "# Page 3" in text
    assert any("Page 2" in warning for warning in result.warnings)


def test_a_missing_chandra_install_fails_the_pdfs_with_an_install_hint(tmp_path: Path, monkeypatch) -> None:
    source = _blank_pdf(tmp_path / "doc.pdf", 1)

    def missing(server=None, model=None):
        raise RuntimeError(chandra_backend.INSTALL_HINT)

    monkeypatch.setattr(chandra_backend, "create_engine", missing)
    batch = convert_files([source], [tmp_path / "doc.md"], ConversionOptions(pdf_backend="chandra"))
    assert not batch.results
    assert "requirements-chandra.txt" in str(batch.failures[0].error)


def test_pdf_backend_is_validated(tmp_path: Path) -> None:
    source = _blank_pdf(tmp_path / "doc.pdf", 1)
    batch = convert_files([source], [tmp_path / "doc.md"], ConversionOptions(pdf_backend="marker"))
    assert "pdf_backend" in str(batch.failures[0].error)


def test_server_engine_points_chandra_at_the_given_endpoint(monkeypatch) -> None:
    import chandra.model
    from chandra.settings import settings

    monkeypatch.setattr(settings, "VLLM_MODEL_NAME", settings.VLLM_MODEL_NAME)
    monkeypatch.setattr(chandra.model, "InferenceManager", lambda method: SimpleNamespace(method=method))
    engine = chandra_backend.create_engine("http://localhost:8080/v1/", "my-chandra")

    assert engine.method == "vllm"
    assert engine.generate_kwargs["vllm_api_base"] == "http://localhost:8080/v1"
    assert settings.VLLM_MODEL_NAME == "my-chandra"


@pytest.fixture()
def entry_point():
    spec = importlib.util.spec_from_file_location("document2md_entry_chandra", SKILL_SCRIPT)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


@pytest.mark.parametrize(
    ("argv", "wanted"),
    [
        (["a.pdf", "--pdf-backend", "chandra"], True),
        (["a.pdf", "--pdf-backend=chandra"], True),
        (["a.pdf", "--pdf-backend", "docling"], False),
        (["a.pdf", "--pdf-backend"], False),
        (["a.pdf"], False),
    ],
)
def test_chandra_is_requested_only_by_the_backend_flag(entry_point, argv, wanted) -> None:
    assert entry_point._wants_chandra(argv) is wanted


def test_chandra_is_installed_only_into_the_skill_environment(entry_point, monkeypatch) -> None:
    calls = []
    monkeypatch.setattr(entry_point.subprocess, "run", lambda *a, **k: calls.append(a))
    monkeypatch.setattr("importlib.util.find_spec", lambda name: None)

    monkeypatch.setattr(entry_point, "_in_skill_environment", lambda: False)
    entry_point._ensure_chandra(["a.pdf", "--pdf-backend", "chandra"])
    assert calls == []

    monkeypatch.setattr(entry_point, "_in_skill_environment", lambda: True)
    entry_point._ensure_chandra(["a.pdf", "--pdf-backend", "chandra", "--no-bootstrap"])
    assert calls == []

    monkeypatch.setattr(
        entry_point.subprocess,
        "run",
        lambda command, **k: calls.append(command) or SimpleNamespace(returncode=0),
    )
    entry_point._ensure_chandra(["a.pdf", "--pdf-backend", "chandra"])
    assert calls and calls[0][-1] == str(entry_point.CHANDRA_REQUIREMENTS)
