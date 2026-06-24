# office2md

`office2md` converts `.docx`, `.pptx`, and `.xlsx` files into readable Markdown without using Pandoc.

The repository is skill-first: the canonical implementation lives inside `skills/office2md`, where Codex can use it directly. The optional CLI package installs the same code.

## Principles

- Preserve source structure instead of guessing intent.
- Keep Word and PowerPoint tables as Markdown tables.
- Keep worksheets separated and represent their used ranges as tables.
- Use real Office metadata, such as Word heading styles and PowerPoint titles, when available.
- Do not hard-wrap prose.
- Extract embedded images to a neighboring assets directory.
- Make lossy decisions explicit and deterministic.

## CLI

```powershell
python -m pip install -e .
office2md document.docx
office2md deck.pptx workbook.xlsx --output converted
office2md workbook.xlsx --values-only --force
```

The CLI writes one Markdown file per input. For a single input, `--output` may be an exact `.md` path. For multiple inputs, it must be a directory.

## Skill

The installable skill is in [`skills/office2md`](skills/office2md). Its bundled script works without installing this repository as a package:

```powershell
python skills/office2md/scripts/convert_office.py input.docx
```

