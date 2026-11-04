#!/usr/bin/env python3
"""Create a Chinese project brief DOCX with python-docx and atomic replace."""
import os
import sys
from pathlib import Path

sys.path.insert(0, os.environ["DSH_OFFICE_HELPERS"])
from atomic import atomic_replace
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Inches, Pt

dest = Path(sys.argv[1] if len(sys.argv) > 1 else "reports/brief.docx")
document = Document()
section = document.sections[0]
section.page_width = Inches(8.5)
section.page_height = Inches(11)
header = section.header.paragraphs[0]
header.text = "项目简报"
footer = section.footer.paragraphs[0]
footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
footer.text = "页码"


def set_run(run, text, size, bold=False):
    run.text = text
    run.bold = bold
    run.font.size = Pt(size)
    run.font.name = "Noto Sans SC"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Noto Sans SC")


title = document.add_paragraph()
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
set_run(title.add_run(), "项目简报", 22, True)
h1 = document.add_paragraph()
h1.style = document.styles["Heading 1"]
set_run(h1.add_run(), "背景", 16, True)
body = document.add_paragraph()
set_run(body.add_run(), "本简报说明范围、进度和下一步。", 12)
h2 = document.add_paragraph()
h2.style = document.styles["Heading 2"]
set_run(h2.add_run(), "关键指标", 14, True)
table = document.add_table(rows=3, cols=3)
for i, row in enumerate([["指标", "本期", "说明"], ["营收", "1200", "万元"], ["毛利", "360", "万元"]]):
    for j, value in enumerate(row):
        table.rows[i].cells[j].text = value
tmp = dest.with_name(f".{dest.name}.buf")
document.save(tmp)
atomic_replace(str(dest), tmp.read_bytes())
tmp.unlink(missing_ok=True)
print("wrote", dest)
print("runtime", os.environ.get("DSH_OFFICE_PYTHON"))
