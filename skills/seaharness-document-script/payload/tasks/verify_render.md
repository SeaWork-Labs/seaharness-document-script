# Task: Verify / render a DOCX (DOCX → PNG)

## Why this exists
DOCX editing tools can "succeed" while the visual output is broken. Always verify by rendering.

## Preferred: use the packaged renderer
This uses a dedicated LibreOffice profile + writable HOME and produces `page-<N>.png` images:

```bash
"$DSH_OFFICE_PYTHON" render_docx.py input.docx --output_dir out
# For debugging LibreOffice failures:
"$DSH_OFFICE_PYTHON" render_docx.py input.docx --output_dir out --verbose
# Optional: also write <input_stem>.pdf to --output_dir (for debugging/archival):
"$DSH_OFFICE_PYTHON" render_docx.py input.docx --output_dir out --emit_pdf
```

## Manual render command (if you need it)
Use a unique LibreOffice profile (permission/locking issues are common in containers).
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

Then rasterize with the bundled PDF renderer instead of `pdftoppm` (poppler is not bundled either):

```bash
"$DSH_OFFICE_PYTHON" - "$OUTDIR/$BASENAME.pdf" "$OUTDIR" <<'PY'
import sys, pypdfium2 as pdfium
pdf_path, out_dir = sys.argv[1], sys.argv[2]
doc = pdfium.PdfDocument(pdf_path)
for i in range(len(doc)):
    page = doc[i]
    page.render(scale=2.0).to_pil().save(f"{out_dir}/page-{i + 1}.png")
    page.close()
doc.close()
PY
```

## Success criteria
- PNGs exist for each page
- Spot-check page count and representative pages

**Note:** LibreOffice sometimes prints scary-looking stderr (e.g., `error : Unknown IO error`) even when output is correct. Treat the conversion as successful if the PNGs exist and look correct (and if you used `--emit_pdf`, the PDF exists and is non-empty).

## What to check in the PNGs
- clipped text (especially headings and table cells)
- overlapping objects
- broken tables (wrapping, misalignment, missing borders)
- unexpected font substitution
- header/footer alignment and page breaks

## Caveats
- **Comments often don’t render** in headless LibreOffice PDFs. Use structural checks for comments.
- **Field codes (page numbers, TOC)** may show placeholder values in some PDF renders. If the user needs proof, re-check in Word or update fields before final render.
- **Multi-section docs** can have different page sizes/orientations; DPI is computed from the first section by default. If some pages look scaled oddly, use `--dpi` to override.

## Delivery checklist
- Final DOCX is clean (no internal citation tokens, no placeholder text)
- Final render looks correct on all pages
- `the Session workspace directory` contains only final outputs (unless user asked for intermediates)


## SeaHarness shipping rule

视觉 PNG 不是本运行时的发运门。若本机有 LibreOffice 且用户要求页图，可跑 render_docx.py。
默认：document_validate + Document Preview。禁止在未渲染时声称「逐页视觉验收已通过」。
