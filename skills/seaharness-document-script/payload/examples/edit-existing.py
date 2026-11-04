#!/usr/bin/env python3
"""Edit an existing DOCX in place: change the title, keep tables, headers, and run formatting."""
import os
import sys
from pathlib import Path

sys.path.insert(0, os.environ["DSH_OFFICE_HELPERS"])
from atomic import atomic_replace, sha256_file
from docx import Document

src = Path(sys.argv[1])
expected = sha256_file(str(src)) if src.exists() else None
document = Document(str(src))
if document.paragraphs and document.paragraphs[0].runs:
    # Assign run.text so w:rPr (font/size/bold/color) stays. paragraph.text wipes runs.
    document.paragraphs[0].runs[0].text = "修订版简报"
    for extra in document.paragraphs[0].runs[1:]:
        extra.text = ""
tmp = src.with_name(f".{src.name}.buf")
document.save(tmp)
atomic_replace(str(src), tmp.read_bytes(), expected)
tmp.unlink(missing_ok=True)
print("updated", src)
