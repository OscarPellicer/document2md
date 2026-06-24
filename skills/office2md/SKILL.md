---
name: office2md
description: Convert Microsoft Office files (.docx, .pptx, and .xlsx) to faithful, readable Markdown using format-aware Python libraries instead of Pandoc. Use when Codex needs an editable or reviewable Markdown representation of Word documents, PowerPoint presentations, or Excel workbooks while preserving tables and source order without inventing semantic structure.
---

# office2md

Convert Office files to Markdown with the bundled format-aware converter. Treat conversion as extraction, not reinterpretation: preserve encoded structure and do not infer headings, fields, checklist controls, or relationships that the source does not explicitly represent.

## Workflow

1. Confirm the input extension is `.docx`, `.pptx`, or `.xlsx`.
2. Run:

```bash
python <skill>/scripts/convert_office.py path/to/input.ext
```

3. For several inputs, pass all paths and an output directory:

```bash
python <skill>/scripts/convert_office.py first.docx second.pptx book.xlsx --output converted
```

4. Inspect the Markdown for structural completeness. Compare the number and order of source tables, slides, and worksheets when fidelity matters.
5. Return the Markdown and its generated assets directory, if one exists.

## Conversion rules

- Preserve paragraphs in source order and do not hard-wrap them.
- Convert every Word or PowerPoint table to a Markdown table.
- Represent merged table cells once; leave repeated grid positions blank because Markdown has no native row/column spans.
- Use Word heading styles as Markdown headings. Do not promote bold paragraphs to headings.
- Represent each PowerPoint slide as a section, using its title placeholder when available.
- Preserve slide reading order using shape position, then shape order.
- Represent each non-empty worksheet as a section and its used range as a Markdown table.
- Preserve formulas by default. Use `--values-only` only when the user wants cached values.
- Extract embedded DOCX/PPTX images beside the Markdown and link them relatively.
- Do not silently repair text split across tables, slides, cells, or paragraphs.

## Options

```text
--output PATH       Exact .md path for one input, or output directory
--assets-dir PATH   Asset directory for one input
--values-only       XLSX: export cached values instead of formulas
--include-hidden    XLSX: include hidden worksheets
--include-notes     PPTX: include speaker notes
--force             Overwrite existing outputs
```

Read [references/format-behavior.md](references/format-behavior.md) when the user asks about fidelity, limitations, or the exact mapping for a format.

