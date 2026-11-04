# Task: Read / review an existing DOCX

## What to review
- Layout: page breaks, margins, clipping/overlap
- Typography: heading hierarchy, font consistency, line spacing
- Tables/figures: alignment, legibility, truncation
- Redlines: do tracked insertions/deletions show up?
- Comments: do they exist (structurally), even if they don’t render?

## Primary method: DOCX → PNG(s) (internally via PDF)

### Preferred: use the packaged renderer
This is the “golden path” because it handles the container-safe LibreOffice profile + HOME automatically and normalizes output names to `page-<N>.png`.

```bash
"$DSH_OFFICE_PYTHON" render_docx.py input.docx --output_dir out
# If debugging LibreOffice:
"$DSH_OFFICE_PYTHON" render_docx.py input.docx --output_dir out --verbose
# Optional: also write <input_stem>.pdf to --output_dir (for debugging/archival):
"$DSH_OFFICE_PYTHON" render_docx.py input.docx --output_dir out --emit_pdf
```

### Manual method (only if debugging)
Use a unique LibreOffice profile + writable HOME (containers are prone to profile permission/locking issues).
`soffice` is not bundled with this runtime — this path only works if the host has it.

```bash
OUTDIR=out
INPUT=input.docx
BASENAME=$(basename "$INPUT" .docx)
LO_PROFILE=.lo_profile_${BASENAME}_$$
mkdir -p "$OUTDIR" "$LO_PROFILE"

HOME="$LO_PROFILE" soffice --headless -env:UserInstallation=file://"$LO_PROFILE" \
  --convert-to pdf --outdir "$OUTDIR" "$INPUT"
```

Rasterize the PDF with the bundled renderer (poppler's `pdftoppm` is not bundled):

```bash
"$DSH_OFFICE_PYTHON" - "$OUTDIR/$BASENAME.pdf" "$OUTDIR" <<'PY'
import sys, pypdfium2 as pdfium
pdf_path, out_dir = sys.argv[1], sys.argv[2]
doc = pdfium.PdfDocument(pdf_path)
for i in range(len(doc)):
    page = doc[i]
    page.render(scale=2.0).to_pil().save(f"{out_dir}/{i + 1}.png")
    page.close()
doc.close()
PY
```

### Success criteria
- Page images exist for each page
- Spot-check page count and representative pages

**Note:** LibreOffice sometimes prints scary-looking stderr (e.g., `error : Unknown IO error`) even when output is correct. Prefer file existence + visual inspection over stderr content.

### Visually inspect every page
Focus on:
- clipped/overlapping text
- tables that wrap unexpectedly
- inconsistent fonts/sizes
- misplaced headers/footers

## Notes on redlines vs comments
- **Tracked changes** (insertions/deletions) often show up in PDF renders.
- **Comments frequently do NOT show up in PDF/image renders** (especially via headless LibreOffice).
  - Rendering is not proof of comments.
  - To verify comments, do a structural check (see `ooxml/comments.md`) or use `pandoc --track-changes=all` to confirm comment markup is present.

## If the doc is huge
Render and inspect key pages first (title, TOC, sections with tables, appendices), then spot-check.
