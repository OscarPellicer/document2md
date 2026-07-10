---
name: document2md
description: Convert PDF and Microsoft Office files (.pdf, .docx, .pptx, and .xlsx) to faithful, readable Markdown using Docling and format-aware Python libraries instead of Pandoc or PyMuPDF. Use when Codex needs an editable or reviewable Markdown representation of PDFs, Word documents, PowerPoint presentations, or Excel workbooks while preserving tables and source order without inventing semantic structure.
---

# document2md

Convert Office files to Markdown with the bundled format-aware converter. Treat conversion as extraction, not reinterpretation: preserve encoded structure and do not infer headings, fields, checklist controls, or relationships that the source does not explicitly represent.

## Workflow

1. Confirm the input extension is `.pdf`, `.docx`, `.pptx`, or `.xlsx`.
2. Choose image handling:
   - Use `--images placeholder` for text-first review, small output, or when images are decorative.
   - Use `--images extract` when figures, diagrams, screenshots, scanned material, or visual evidence matter. Return the generated asset folder with the Markdown.
   - Use `--images embed` only for a self-contained PDF Markdown file when its larger size is acceptable.
3. Run:

```bash
python <skill>/scripts/convert_document.py path/to/input.ext --images placeholder
```

   When the `document2md` repository has its own virtual environment, prefer that interpreter even if the current workspace is elsewhere: `C:\Users\Oscar\document2md\.venv\Scripts\python.exe` on this Windows machine, or `<document2md-repo>/.venv/bin/python` on Unix. If the repo-local venv is unavailable, use an environment with `scripts/requirements.txt` installed.
   PDFs are read through an ASCII-named in-memory stream, so Unicode characters in the source path do not reach Docling's native backend. PDFs are processed in ordered 10-page batches by default to cap memory use.

4. For several inputs, pass all paths and an output directory:

```bash
python <skill>/scripts/convert_document.py first.docx second.pptx book.xlsx --output converted
```

   The CLI batches all PDF inputs through one shared Docling `DocumentConverter`. Do not invoke the script separately in a loop for multiple PDFs.

5. Inspect the Markdown for structural completeness. Compare the number and order of source tables, slides, and worksheets when fidelity matters.
6. Return the Markdown and its generated assets directory when `--images extract` was selected.

## Conversion rules

- Preserve paragraphs in source order and do not hard-wrap them.
- Convert PDFs with Docling's standard pipeline, including OCR and layout-aware table reconstruction when needed.
- Convert every Word or PowerPoint table to a Markdown table.
- Represent merged table cells once; leave repeated grid positions blank because Markdown has no native row/column spans.
- When a Word table is nested inside another table cell, render the nested table as compact inline text inside that cell. Markdown cannot represent true nested tables, so preserve source order and cell content instead of emitting escaped table syntax that looks like a broken table.
- Use Word heading styles as Markdown headings. Do not promote bold paragraphs to headings.
- Represent each PowerPoint slide as a section, using its title placeholder when available.
- Preserve slide reading order using shape position, then shape order.
- Represent each non-empty worksheet as a section and its used range as a Markdown table.
- Preserve formulas by default. Use `--values-only` only when the user wants cached values.
- Extract PDF/DOCX/PPTX images beside the Markdown and link them relatively only when `--images extract` is selected.
- Do not silently repair text split across tables, slides, cells, or paragraphs.

## Options

```text
--output PATH       Exact .md path for one input, or output directory
--assets-dir PATH   Asset directory for one input
--values-only       XLSX: export cached values instead of formulas
--include-hidden    XLSX: include hidden worksheets
--include-notes     PPTX: include speaker notes
--images MODE       placeholder (default), extract, or embed (PDF only)
--pdf-pages-per-batch N
                    PDF pages processed at once (default: 10; 0 disables batching)
--force             Overwrite existing outputs
```

## PDF troubleshooting

- Keep the default `--pdf-pages-per-batch 10` for long or layout-heavy PDFs.
- If memory is unusually constrained, retry with `--pdf-pages-per-batch 5` or `1`.
- Use `--pdf-pages-per-batch 0` only when whole-document processing is specifically needed and sufficient memory is available.
- Do not manually rename or copy a PDF solely because its path contains spaces, accents, or other Unicode characters; the converter handles this internally.

Read [references/format-behavior.md](references/format-behavior.md) when the user asks about fidelity, limitations, or the exact mapping for a format.
