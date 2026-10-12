# document2md

`document2md` converts `.pdf`, `.docx`, `.pptx`, `.xlsx`, `.odt`, `.ods`, and `.odp` files into readable Markdown without using Pandoc. PDF conversion uses Docling rather than PyMuPDF; OpenDocument files are read straight from their XML.

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
- Link extracted images relatively, so a Markdown file and its assets folder can be moved together.
- Never report success for a file that was not written.
- Let one unreadable input fail on its own without discarding the rest of a batch.

## CLI

```powershell
python -m pip install -e .
document2md document.pdf --images extract
document2md scanned.pdf --ocr force
document2md born-digital.pdf --ocr off
document2md long-document.pdf --pdf-pages-per-batch 5
document2md paper.pdf --formulas
document2md scan.pdf --pdf-backend chandra
document2md scan.pdf --pdf-backend chandra --chandra-server http://127.0.0.1:8080/v1
document2md document.docx
document2md deck.pptx workbook.xlsx --output converted
document2md workbook.xlsx --values-only --force
document2md form.odt notes.odp
```

The CLI writes one Markdown file per input. For a single input, `--output` may be an exact `.md` path. For multiple inputs, it must be a directory.

A file that cannot be converted is reported on stderr and the batch continues; the exit code is `1` when any input failed, `2` when the whole invocation was rejected. Use `--strict` to stop at the first failure instead.

`--ocr` controls the PDF OCR stage: `auto` (default) runs OCR on bitmap areas, `off` skips it entirely and is much faster on born-digital PDFs, and `force` OCRs every page, which is what scans need.

`--pdf-backend chandra` converts PDFs with [Chandra](https://github.com/datalab-to/chandra), a vision-language OCR model, instead of Docling. It is opt-in: the `chandra-ocr` package is installed into the skill environment on first use, the model weights are about 10.6 GB, and on an Apple M5 Pro it takes 20-100 s per page locally against 1-4 s for Docling. It reads scans and outlined text far better, writes equations as LaTeX, and keeps citations as superscripts. Being a language model, it can also alter names and identifiers. Use `--chandra-server URL` to send pages to a vLLM or `mlx_vlm.server` endpoint instead of loading the model in-process. See the skill's `SKILL.md` for when to choose it.

`--formulas` decodes PDF equations to LaTeX with Docling's formula model. It is off by default because it takes minutes per page on CPU.

PDF inputs are passed to Docling through an ASCII-named memory stream, so paths containing spaces, accents, or other Unicode characters are supported. Long PDFs are processed in ordered 10-page ranges by default to limit peak memory use. Use `--pdf-pages-per-batch 5` (or `1`) on a constrained machine, and `--pdf-pages-per-batch 0` to disable batching.

The script creates and reuses its own `.venv` (see [Skill](#skill)), so for local development the gitignored `.venv` already contains all runtime and test dependencies:

```powershell
.\.venv\Scripts\document2md.exe input.pdf --images extract
.\.venv\Scripts\python.exe -m pytest
```

## Skill

The installable skill is in [`skills/document2md`](skills/document2md). Its bundled script works without installing this repository as a package:

```powershell
python skills/document2md/scripts/convert_document.py input.docx
```

On its first run the script creates a virtual environment at `skills/document2md/.venv`, installs `scripts/requirements.txt` into it, and re-executes itself with that interpreter. Every later run finds the same environment, so callers never have to locate or activate it, and no caller needs to hardcode an interpreter path. Pass `--no-bootstrap` to stay in the current interpreter.
