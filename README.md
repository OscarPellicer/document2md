# document2md

`document2md` converts `.pdf`, `.docx`, `.pptx`, and `.xlsx` files into readable Markdown without using Pandoc. PDF conversion uses Docling rather than PyMuPDF.

The repository is skill-first: the canonical implementation lives inside `skills/document2md`, where Codex can use it directly. The optional CLI package installs the same code.

## Principles

- Preserve source structure instead of guessing intent.
- Keep Word and PowerPoint tables as Markdown tables.
- Keep worksheets separated and represent their used ranges as tables.
- Use real Office metadata, such as Word heading styles and PowerPoint titles, when available.
- Do not hard-wrap prose.
- Extract embedded images to a neighboring assets directory.
- Let callers choose placeholder, extracted, or PDF-embedded image handling.
- Make lossy decisions explicit and deterministic.

## CLI

```powershell
python -m pip install -e .
document2md document.pdf --images extract
document2md document.docx
document2md deck.pptx workbook.xlsx --output converted
document2md workbook.xlsx --values-only --force
```

The CLI writes one Markdown file per input. For a single input, `--output` may be an exact `.md` path. For multiple inputs, it must be a directory.

For local development, the gitignored `.venv` contains all runtime and test dependencies:

```powershell
.\.venv\Scripts\document2md.exe input.pdf --images extract
.\.venv\Scripts\python.exe -m pytest
```

## Skill

The installable skill is in [`skills/document2md`](skills/document2md). Its bundled script works without installing this repository as a package:

```powershell
python skills/document2md/scripts/convert_document.py input.docx
```
