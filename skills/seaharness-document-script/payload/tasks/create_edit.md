# Task: Create / edit a DOCX

## Default tool: python-docx
Use `python-docx` for:
- paragraphs/runs
- built-in styles (Title / Heading 1 / Heading 2)
- tables (structure + cell text + basic formatting)
- simple headers/footers and margins

When editing an existing DOCX, change wording only. Keep the original font, size, bold/italic, color, underline, paragraph style, alignment, and spacing unless the user asked to restyle. Replace `run.text` (or the `<w:t>` node). Do **not** assign `paragraph.text` — python-docx rebuilds the paragraph as a single unstyled run. For tracked changes, copy `<w:rPr>` onto `<w:del>` / `<w:ins>` runs (`scripts/add_tracked_replacements.py`).

## Practical python-docx gotchas

### 1) Header/footer tables require a width
When adding tables to headers/footers, `add_table` requires an explicit width:

```python
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

section = doc.sections[0]
footer = section.footer
table = footer.add_table(rows=1, cols=3, width=Inches(6.5))
# Align text inside each cell
table.rows[0].cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
```

### 2) Fonts can require setting both `run.font.name` and `w:rFonts`
Some renderers/Word builds don’t respect only `run.font.name`:

```python
from docx.oxml.ns import qn

run.font.name = "Gill Sans"
run._element.rPr.rFonts.set(qn("w:ascii"), "Gill Sans")
run._element.rPr.rFonts.set(qn("w:hAnsi"), "Gill Sans")
```

### 3) “Clear header paragraph” isn’t always one call
If you need to replace an existing header paragraph, remove runs (or replace the paragraph XML). Avoid assuming a `clear()` method exists.

### 4) Tracked changes and comments are not first-class
If the user requests *real* tracked changes or *real* Word comments, plan for OOXML patching (see `ooxml/`).

## After every meaningful batch of edits: render and review
Use the loop from `tasks/verify_render.md` (DOCX → PNG) to avoid shipping layout defects. (Internally the renderer uses a PDF step; `--emit_pdf` can persist it if needed.)

## Output hygiene
Keep `the Session workspace directory` clean: deliverables only unless the user asks for intermediate render artifacts.
Do not stamp SeaHarness / tool branding into headers, footers, or body. A footer, if present, is a page number or the user's organization — never「本文档由 SeaHarness 自动生成」.
