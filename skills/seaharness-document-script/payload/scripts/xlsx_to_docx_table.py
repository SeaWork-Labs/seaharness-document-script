#!/usr/bin/env python3
"""Convert an XLSX worksheet to a simple DOCX table.

Scope
-----
This is a "minimum useful" converter intended for reports:
- Reads values from a sheet
- Creates a DOCX with a grid table
- Applies a few common-sense formats (header bold, basic alignment)
- Uses heuristic column widths based on max content length

Limits
------
- It does NOT reproduce Excel formulas, merged cells, conditional formatting, charts, or complex number formats.
- Number formatting is best-effort. If you need exact appearance, consider exporting the range as an image instead.

The workbook is read with the Python standard library (zipfile + XML) rather than
openpyxl, because the Office runtime ships only python-docx, lxml, Pillow,
reportlab, pypdf, and pdfplumber.
"""

from __future__ import annotations

import argparse
import datetime as _dt
import os
import re
import sys
import zipfile
from pathlib import Path
from typing import Iterator
from xml.etree import ElementTree as ET

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches

_NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"

# Built-in Excel number formats that render as dates/times.
_DATE_NUMFMT_IDS = {14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47}

# Built-in format codes, so cell.number_format carries the same text openpyxl
# reports. Only ids whose code affects this converter's output are listed;
# everything else is left empty and rendered as a plain number.
_BUILTIN_NUMFMTS = {
    1: "0",
    2: "0.00",
    3: "#,##0",
    4: "#,##0.00",
    5: '"$"#,##0_);("$"#,##0)',
    6: '"$"#,##0_);[Red]("$"#,##0)',
    7: '"$"#,##0.00_);("$"#,##0.00)',
    8: '"$"#,##0.00_);[Red]("$"#,##0.00)',
    9: "0%",
    10: "0.00%",
    11: "0.00E+00",
    37: "#,##0 ;(#,##0)",
    38: "#,##0 ;[Red](#,##0)",
    39: "#,##0.00;(#,##0.00)",
    40: "#,##0.00;[Red](#,##0.00)",
    48: "##0.0E+0",
    49: "@",
}

_EPOCH_1900 = _dt.datetime(1899, 12, 30)
_EPOCH_1904 = _dt.datetime(1904, 1, 1)


def _q(tag: str, ns: str = _NS_MAIN) -> str:
    return f"{{{ns}}}{tag}"


class _Alignment:
    def __init__(self, horizontal: str | None = None) -> None:
        self.horizontal = horizontal


class _Cell:
    __slots__ = ("value", "row", "column", "number_format", "alignment")

    def __init__(self, value, row: int, column: int, number_format: str, horizontal: str | None) -> None:
        self.value = value
        self.row = row
        self.column = column
        self.number_format = number_format
        self.alignment = _Alignment(horizontal)


class _Worksheet:
    def __init__(self, title: str, cells: dict[tuple[int, int], _Cell], max_row: int, max_col: int) -> None:
        self.title = title
        self._cells = cells
        self.max_row = max_row
        self.max_column = max_col

    def cell(self, row: int, column: int) -> _Cell:
        found = self._cells.get((row, column))
        if found is not None:
            return found
        return _Cell(None, row, column, "", None)

    def iter_rows(self) -> Iterator[tuple[_Cell, ...]]:
        for r in range(1, self.max_row + 1):
            yield tuple(self.cell(r, c) for c in range(1, self.max_column + 1))


class _Workbook:
    def __init__(self, sheets: dict[str, _Worksheet], order: list[str]) -> None:
        self._sheets = sheets
        self._order = order
        self.active = sheets[order[0]] if order else None

    def __getitem__(self, name: str) -> _Worksheet:
        return self._sheets[name]

    def __contains__(self, name: object) -> bool:
        return name in self._sheets

    @property
    def sheetnames(self) -> list[str]:
        return list(self._order)


def _text_of(node) -> str:
    """Concatenate every <t> descendant, which covers <si><t> and <si><r><t>."""
    return "".join(t.text or "" for t in node.iter(_q("t")))


def _column_index(ref: str) -> int:
    letters = "".join(ch for ch in ref if ch.isalpha())
    index = 0
    for ch in letters.upper():
        index = index * 26 + (ord(ch) - 64)
    return index


def _row_index(ref: str) -> int:
    digits = "".join(ch for ch in ref if ch.isdigit())
    return int(digits) if digits else 0


def _parse_ref(ref: str) -> tuple[int, int]:
    return _row_index(ref), _column_index(ref)


def _is_date_format(format_id: int, custom: dict[int, str]) -> bool:
    if format_id in _DATE_NUMFMT_IDS:
        return True
    code = custom.get(format_id)
    if not code:
        return False
    # Drop quoted literals, bracketed sections (colors/locales), and escapes so
    # only real format tokens remain.
    cleaned = re.sub(r'"[^"]*"', "", code)
    cleaned = re.sub(r"\[[^\]]*\]", "", cleaned)
    cleaned = re.sub(r"\\.", "", cleaned)
    lowered = cleaned.lower()
    return any(ch in lowered for ch in "ydhs")


def _serial_to_datetime(serial: float, date1904: bool) -> _dt.datetime | _dt.date:
    epoch = _EPOCH_1904 if date1904 else _EPOCH_1900
    value = epoch + _dt.timedelta(days=serial)
    if value.time() == _dt.time(0, 0):
        return value.date()
    return value


def _read_styles(zf: zipfile.ZipFile) -> tuple[list[tuple[int, str | None]], dict[int, str]]:
    """Return per-cellXfs (number_format_id, horizontal) plus custom format codes."""
    try:
        raw = zf.read("xl/styles.xml")
    except KeyError:
        return [], {}

    root = ET.fromstring(raw)

    custom: dict[int, str] = {}
    numfmts = root.find(_q("numFmts"))
    if numfmts is not None:
        for numfmt in numfmts.findall(_q("numFmt")):
            try:
                custom[int(numfmt.get("numFmtId"))] = numfmt.get("formatCode") or ""
            except (TypeError, ValueError):
                continue

    formats: list[tuple[int, str | None]] = []
    cellxfs = root.find(_q("cellXfs"))
    if cellxfs is not None:
        for xf in cellxfs.findall(_q("xf")):
            try:
                format_id = int(xf.get("numFmtId") or 0)
            except ValueError:
                format_id = 0
            alignment = xf.find(_q("alignment"))
            horizontal = alignment.get("horizontal") if alignment is not None else None
            formats.append((format_id, horizontal))
    return formats, custom


def _read_shared_strings(zf: zipfile.ZipFile) -> list[str]:
    try:
        raw = zf.read("xl/sharedStrings.xml")
    except KeyError:
        return []
    root = ET.fromstring(raw)
    return [_text_of(si) for si in root.findall(_q("si"))]


def _sheet_targets(zf: zipfile.ZipFile) -> dict[str, str]:
    """Map workbook relationship id -> worksheet part name."""
    try:
        raw = zf.read("xl/_rels/workbook.xml.rels")
    except KeyError:
        return {}
    root = ET.fromstring(raw)
    targets: dict[str, str] = {}
    for rel in root.findall(f"{{{_NS_PKG_REL}}}Relationship"):
        rid = rel.get("Id")
        target = rel.get("Target")
        if not rid or not target:
            continue
        if target.startswith("/"):
            target = target.lstrip("/")
        elif not target.startswith("xl/"):
            target = "xl/" + target
        targets[rid] = os.path.normpath(target).replace(os.sep, "/")
    return targets


def _read_sheet(
    zf: zipfile.ZipFile,
    part: str,
    title: str,
    shared: list[str],
    formats: list[tuple[int, str | None]],
    custom: dict[int, str],
    date1904: bool,
) -> _Worksheet:
    root = ET.fromstring(zf.read(part))
    cells: dict[tuple[int, int], _Cell] = {}
    max_row = 0
    max_col = 0

    sheet_data = root.find(_q("sheetData"))
    if sheet_data is None:
        return _Worksheet(title, cells, 0, 0)

    for row_node in sheet_data.findall(_q("row")):
        for cell_node in row_node.findall(_q("c")):
            ref = cell_node.get("r") or ""
            if ref:
                row_idx, col_idx = _parse_ref(ref)
            else:
                # Sparse writers may omit r=; fall back to declared row and order.
                row_idx = int(row_node.get("r") or 0)
                col_idx = len(cells) + 1
            if row_idx <= 0 or col_idx <= 0:
                continue

            cell_type = cell_node.get("t") or "n"
            raw_value = cell_node.findtext(_q("v"))

            value = None
            if cell_type == "inlineStr":
                inline = cell_node.find(_q("is"))
                value = _text_of(inline) if inline is not None else None
            elif cell_type == "s":
                if raw_value is not None and raw_value.strip():
                    try:
                        value = shared[int(raw_value)]
                    except (ValueError, IndexError):
                        value = None
            elif cell_type == "b":
                value = None if raw_value is None else raw_value.strip() == "1"
            elif cell_type == "e":
                value = None
            elif cell_type == "str":
                value = raw_value
            elif cell_type == "d":
                if raw_value:
                    try:
                        value = _dt.datetime.fromisoformat(raw_value)
                    except ValueError:
                        value = raw_value
            elif raw_value is not None and raw_value.strip():
                try:
                    number = float(raw_value)
                except ValueError:
                    value = raw_value
                else:
                    style_index = cell_node.get("s")
                    try:
                        fmt_id, _ = formats[int(style_index)] if style_index else (0, None)
                    except (IndexError, ValueError):
                        fmt_id = 0
                    if _is_date_format(fmt_id, custom):
                        try:
                            value = _serial_to_datetime(number, date1904)
                        except (OverflowError, ValueError):
                            value = number
                    elif number.is_integer():
                        value = int(number)
                    else:
                        value = number

            style_index = cell_node.get("s")
            try:
                fmt_id, horizontal = formats[int(style_index)] if style_index else (0, None)
            except (IndexError, ValueError):
                fmt_id, horizontal = 0, None
            number_format = custom.get(fmt_id) or _BUILTIN_NUMFMTS.get(fmt_id, "")

            cells[(row_idx, col_idx)] = _Cell(value, row_idx, col_idx, number_format, horizontal)
            max_row = max(max_row, row_idx)
            max_col = max(max_col, col_idx)

    return _Worksheet(title, cells, max_row, max_col)


def load_workbook(path: Path) -> _Workbook:
    """Minimal read-only workbook loader (values only, no formulas)."""
    with zipfile.ZipFile(path) as zf:
        formats, custom = _read_styles(zf)
        shared = _read_shared_strings(zf)
        targets = _sheet_targets(zf)

        try:
            workbook_root = ET.fromstring(zf.read("xl/workbook.xml"))
        except KeyError as exc:
            raise RuntimeError(f"Not an XLSX workbook (missing xl/workbook.xml): {path}") from exc

        workbook_pr = workbook_root.find(_q("workbookPr"))
        date1904 = (workbook_pr is not None and (workbook_pr.get("date1904") or "").lower() in {"1", "true"})

        sheets: dict[str, _Worksheet] = {}
        order: list[str] = []
        sheets_node = workbook_root.find(_q("sheets"))
        if sheets_node is None:
            return _Workbook(sheets, order)

        for index, sheet_node in enumerate(sheets_node.findall(_q("sheet"))):
            title = sheet_node.get("name") or f"Sheet{index + 1}"
            rid = sheet_node.get(f"{{{_NS_REL}}}id")
            part = targets.get(rid or "")
            if not part or part not in zf.namelist():
                # Fall back to the conventional part name for the sheet position.
                candidate = f"xl/worksheets/sheet{index + 1}.xml"
                if candidate not in zf.namelist():
                    continue
                part = candidate
            sheets[title] = _read_sheet(zf, part, title, shared, formats, custom, date1904)
            order.append(title)

        return _Workbook(sheets, order)


def _mark_row_as_header(row) -> None:
    """Mark a table row as a repeating header row (w:tblHeader).

    This improves accessibility and ensures header rows repeat when the table
    spans multiple pages.
    """

    tr = row._tr  # python-docx internal
    trPr = tr.get_or_add_trPr()
    # Avoid duplicating if re-run
    if trPr.find(qn("w:tblHeader")) is None:
        trPr.append(OxmlElement("w:tblHeader"))

def _is_empty(v) -> bool:
    if v is None:
        return True
    if isinstance(v, str) and not v.strip():
        return True
    return False

def _best_effort_format(cell) -> str:
    v = cell.value
    if v is None:
        return ""
    if isinstance(v, (bool, int)):
        return str(v)
    if isinstance(v, float):
        fmt = (cell.number_format or "").lower()
        if "%" in fmt:
            return f"{v * 100:.2f}%"
        if "$" in fmt:
            return f"${v:,.2f}"
        return f"{v:g}"
    if isinstance(v, (_dt.datetime, _dt.date)):
        try:
            return v.isoformat()
        except Exception:
            return str(v)
    return str(v)

def _used_bounds(ws):
    # Compute a tight-ish used range; ws.max_row/col may include trailing empties.
    max_r = 0
    max_c = 0
    for row in ws.iter_rows():
        for cell in row:
            if not _is_empty(cell.value):
                max_r = max(max_r, cell.row)
                max_c = max(max_c, cell.column)
    return max_r, max_c

def main() -> None:
    ap = argparse.ArgumentParser(description="Convert an XLSX sheet to a DOCX table")
    ap.add_argument("xlsx", type=Path)
    # Keep the original positional output path for backwards compatibility, but also
    # support an --out alias (agents often assume this flag exists).
    ap.add_argument("out_docx", type=Path, nargs="?", default=None)
    ap.add_argument("--out", type=Path, default=None, help="Output DOCX path (alias)")
    ap.add_argument("--sheet", default=None, help="Sheet name (default: active)")
    ap.add_argument("--header_rows", type=int, default=1, help="How many top rows to bold")
    ap.add_argument("--title", default=None, help="Optional title paragraph above the table")
    args = ap.parse_args()

    out_path = args.out or args.out_docx
    if out_path is None:
        raise RuntimeError("Missing output path. Provide a positional out_docx or --out.")

    if not args.xlsx.exists():
        raise FileNotFoundError(args.xlsx)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    wb = load_workbook(args.xlsx)
    if args.sheet:
        ws = wb[args.sheet]
    elif wb.active is not None:
        ws = wb.active
    else:
        raise RuntimeError("Workbook contains no readable worksheets")

    max_r, max_c = _used_bounds(ws)
    if max_r == 0 or max_c == 0:
        raise RuntimeError("Sheet appears empty")

    # Gather formatted strings for width heuristics
    data = []
    col_max_len = [0] * max_c
    for r in range(1, max_r + 1):
        row_vals = []
        for c in range(1, max_c + 1):
            cell = ws.cell(row=r, column=c)
            s = _best_effort_format(cell)
            row_vals.append(s)
            col_max_len[c - 1] = max(col_max_len[c - 1], len(s))
        data.append(row_vals)

    doc = Document()
    if args.title:
        doc.add_paragraph(args.title)

    table = doc.add_table(rows=max_r, cols=max_c)
    table.style = "Table Grid"
    table.autofit = False

    # Column widths (heuristic): ~0.12in per char, clamped
    col_widths = []
    for ml in col_max_len:
        w = max(0.8, min(3.0, 0.12 * max(ml, 4)))
        col_widths.append(w)

    for r in range(max_r):
        for c in range(max_c):
            cell = ws.cell(row=r + 1, column=c + 1)
            out = table.cell(r, c)
            out.width = Inches(col_widths[c])
            p = out.paragraphs[0]
            p.text = data[r][c]

            # Alignment best-effort
            horiz = getattr(cell.alignment, "horizontal", None)
            if horiz in {"right", "center", "left"}:
                p.alignment = {
                    "left": WD_ALIGN_PARAGRAPH.LEFT,
                    "center": WD_ALIGN_PARAGRAPH.CENTER,
                    "right": WD_ALIGN_PARAGRAPH.RIGHT,
                }[horiz]

            # Bold header rows
            if r < args.header_rows and p.runs:
                for run in p.runs:
                    run.bold = True

        # Mark header rows in OOXML for a11y + repeating headers
        if r < args.header_rows:
            _mark_row_as_header(table.rows[r])

    doc.save(str(out_path))
    print(f"[OK] Wrote {out_path} (rows={max_r}, cols={max_c}, sheet={ws.title!r})")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        if os.environ.get("DOCS_DEBUG") == "1":
            raise
        print(f"[ERROR] {e}", file=sys.stderr)
        raise SystemExit(2)
