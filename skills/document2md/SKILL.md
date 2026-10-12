---
name: document2md
description: Convert PDF, Microsoft Office and OpenDocument files (.pdf, .docx, .pptx, .xlsx, .odt, .ods, .odp) to faithful, readable Markdown using Docling (or, opt-in, the Chandra OCR model for scans) and format-aware Python libraries instead of Pandoc or PyMuPDF. Use when Codex needs an editable or reviewable Markdown representation of PDFs, Word documents, PowerPoint presentations, Excel workbooks, or OpenDocument files while preserving tables and source order without inventing semantic structure.
---

# document2md

Convert documents to Markdown with the bundled format-aware converter. Treat conversion as extraction, not reinterpretation: preserve encoded structure and do not infer headings, fields, checklist controls, or relationships that the source does not explicitly represent.

## Workflow

1. Confirm the input extension is `.pdf`, `.docx`, `.pptx`, `.xlsx`, `.odt`, `.ods`, or `.odp`.
2. Choose image handling:
   - Use `--images placeholder` for text-first review, small output, or when images are decorative.
   - Use `--images extract` when figures, diagrams, screenshots, scanned material, or visual evidence matter. Return the generated asset folder with the Markdown.
   - Use `--images embed` only for a self-contained PDF Markdown file when its larger size is acceptable.
3. Run:

```bash
python <skill>/scripts/convert_document.py path/to/input.ext --images placeholder
```

   Run it with whatever `python`/`python3` is on PATH, even an older one. On its
   first run the script finds a Python 3.11+ (on PATH, or through `uv python find`),
   creates `<skill>/.venv` with it, installs `scripts/requirements.txt`, and
   re-executes itself there; every later run reuses that environment. An
   environment built by an older Python is rebuilt automatically. Do not hunt for
   an interpreter, hardcode one, or activate anything by hand. Pass
   `--no-bootstrap` to stay in the current interpreter when its dependencies are
   already installed.
   PDFs are read through an ASCII-named in-memory stream, so Unicode characters in the source path do not reach Docling's native backend. PDFs are processed in ordered 10-page batches by default to cap memory use.

4. For several inputs, pass all paths and an output directory:

```bash
python <skill>/scripts/convert_document.py first.docx second.pptx book.xlsx --output converted
```

   The CLI batches all PDF inputs through one shared Docling `DocumentConverter` (or one loaded Chandra model). Do not invoke the script separately in a loop for multiple PDFs.
   An input that fails is reported on stderr and the batch continues, so a single unreadable file does not cost the rest of the run. Exit code `1` means some inputs failed while others were converted; `2` means the invocation itself was rejected. Read stderr and report which inputs failed rather than assuming a whole-batch success.

5. Choose OCR handling for PDFs with `--ocr`:
   - `auto` (default) OCRs bitmap areas only.
   - `off` skips OCR entirely. Much faster, and correct for born-digital PDFs.
   - `force` OCRs every page. Use it for scans and for pages whose text layer is missing or wrong.
6. Inspect the Markdown for structural completeness. Compare the number and order of source tables, slides, and worksheets when fidelity matters.
7. Pass `--formulas` only when equations matter: it decodes them to LaTeX with Docling's formula model, which is slow on CPU (minutes per page). Without it each equation is left as `<!-- formula-not-decoded -->`.
8. Keep the default `--pdf-backend docling` unless the PDF has no usable text layer. Switch to `--pdf-backend chandra` (see [Chandra backend](#chandra-backend)) when the user accepts a much slower run for:
   - scans, photos, faxes or handwriting, where Docling's OCR garbles words;
   - slides or posters whose text is outlined or rasterized;
   - equation-heavy pages, where it writes LaTeX for display and inline math.
   Do not use it for born-digital PDFs whose text layer is good: Docling copies that text exactly, while Chandra re-reads the page image and can alter identifiers.
9. Return the Markdown and its generated assets directory when `--images extract` was selected.

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
- Represent each OpenDocument heading, list, paragraph, and table in stored order; `.ods` sheets and `.odp` slides become sections like their Office equivalents.
- Extract PDF/DOCX/PPTX/ODF images beside the Markdown and link them relatively only when `--images extract` is selected. Image links are always relative to the Markdown file, so the file and its assets folder can be moved together.
- Do not silently repair text split across tables, slides, cells, or paragraphs.

## Options

```text
--output PATH       Exact .md path for one input, or output directory
--assets-dir PATH   Asset directory for one input
--values-only       XLSX: export cached values instead of formulas
--include-hidden    XLSX/ODS: include hidden worksheets
--include-notes     PPTX: include speaker notes
--images MODE       placeholder (default), extract, or embed (PDF only)
--ocr MODE          PDF OCR: auto (default), off, or force
--formulas          PDF: decode equations to LaTeX (slow)
--pdf-backend NAME  PDF engine: docling (default) or chandra (optional OCR model)
--chandra-server URL
                    Chandra: OpenAI-compatible server hosting the model
                    (default: $DOCUMENT2MD_CHANDRA_SERVER, else run locally)
--chandra-model NAME
                    Chandra: served model name, or local checkpoint
--strict            Stop at the first failing input instead of converting the rest
--no-bootstrap      Do not create or re-exec into the skill's own virtual environment
--pdf-pages-per-batch N
                    PDF pages processed at once (default: 10; 0 disables batching)
--force             Overwrite existing outputs
```

## PDF troubleshooting

- Keep the default `--pdf-pages-per-batch 10` for long or layout-heavy PDFs.
- If memory is unusually constrained, retry with `--pdf-pages-per-batch 5` or `1`.
- Use `--pdf-pages-per-batch 0` only when whole-document processing is specifically needed and sufficient memory is available.
- Do not manually rename or copy a PDF solely because its path contains spaces, accents, or other Unicode characters; the converter handles this internally.
- Use `--ocr off` for born-digital PDFs. It removes the OCR stage, which dominates runtime, and silences `RapidOCR returned empty result!` notices that do not indicate lost text.
- Words run together (e.g. `Weclarifiedthe...`) come from Docling's parser on tightly set fonts. The converter re-reads each text block with PDFium and keeps that reading when it has the same letters and digits, so this should now be rare; table cells are not re-read.
- Docling reconstructs reading order and tables statistically. Heading order and table cells can be wrong on dense layouts, and signature blocks, footers, and CSV stamps are not carried over. Check the original before relying on such a passage; do not repair it silently in the Markdown.

## Chandra backend

`--pdf-backend chandra` runs Datalab's Chandra OCR 2 (`datalab-to/chandra-ocr-2`, 5.3B parameters), a vision-language model that reads each page as an image. It is opt-in:

- First use installs `scripts/requirements-chandra.txt` into the skill's `.venv` (about 15 MB, since torch and transformers already come with Docling), and the first local run downloads about 10.6 GB of weights into the Hugging Face cache. Expect a one-off delay.
- Locally it runs through transformers on CUDA, Apple MPS or CPU. On an Apple M5 Pro this took about 20-100 s per page (19 tokens/s, about 11 GB of RAM), against about 1-4 s per page for Docling. A 30-page paper takes 20-40 minutes. Warn the user before starting a long PDF, and convert a page subset first when unsure.
- `--chandra-server URL` (or `DOCUMENT2MD_CHANDRA_SERVER`) sends pages to an OpenAI-compatible server instead, such as vLLM on a GPU or `mlx_vlm.server` with an MLX conversion on Apple Silicon (about 1.5-2x faster than local MPS). `--chandra-model` names the served model when the server hosts several. Never point it at a paid API without the user's approval.
- `--ocr` and `--formulas` do not apply: every page is read as an image, and equations are always LaTeX.
- Strengths: clean OCR of scans, LaTeX equations, citations kept as `<sup>`, tables with spans, figure captions kept apart from the figure, and no HTML entities.
- Risks: it is a language model. It can "correct" text it reads, for example a misspelled word, a rare surname, or gene symbols in a dense list (`DGCR8` written as `DGC8R`). Check names, codes and numbers against the source before relying on them. Page headers and footers are dropped, and a running header sometimes survives on one page and not another.
- Pictures become crops of the page (`--images extract`/`embed`) or `<!-- image -->` placeholders. The model's own descriptions of figures, chart values and diagrams are discarded because they are not source text.
- The model weights use a modified OpenRAIL-M licence: free for research, personal use and organisations under $2M in funding or revenue; other commercial use needs a licence from Datalab.

Read [references/format-behavior.md](references/format-behavior.md) when the user asks about fidelity, limitations, or the exact mapping for a format.
