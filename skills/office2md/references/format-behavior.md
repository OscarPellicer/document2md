# Format behavior

## DOCX

- Body paragraphs and tables are emitted in document order.
- Real Word heading styles map to Markdown headings.
- Character-level bold and italic formatting is retained where practical.
- Tables use the physical Word grid. Merged-cell text appears once and duplicate grid positions remain empty.
- Inline and floating drawing extraction is best-effort. Images with package relationships are copied to the assets directory.
- Headers, footers, comments, tracked-change history, text boxes, equations, and exact pagination are not reconstructed.

## PPTX

- Each slide becomes a level-one section.
- The title placeholder supplies the slide heading when present.
- Text boxes become paragraphs; tables become Markdown tables.
- Shapes are ordered top-to-bottom, then left-to-right, with original shape order as a tie-breaker.
- Pictures are extracted and linked.
- Charts, SmartArt, transitions, animations, and visual positioning are not reproduced.
- Speaker notes are optional.

## XLSX

- Each included worksheet becomes a level-one section.
- The rectangular used range becomes one Markdown table.
- Empty trailing rows and columns are trimmed.
- Formulas are retained by default.
- `--values-only` requests cached formula results; results may be empty when the workbook was not recalculated before saving.
- Merged-cell content appears in the top-left cell and other positions remain empty.
- Styling, charts, conditional formatting, drawings, comments, and formulas' calculated display formatting are not reproduced.

## Why this differs from Pandoc

Pandoc is a broad document converter. Office files often use tables and drawing objects for layout, which can lead to large HTML fragments in Markdown output. This converter uses Office-specific object models and deliberately emits a smaller Markdown vocabulary. It favors predictable structural fidelity over visual imitation or semantic guessing.

