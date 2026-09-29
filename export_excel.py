"""
export_excel.py  —  CRD Chargeability Skill  (scripts/)
--------------------------------------------------------
Embeds Tool Assessment, Tool Reason, and Scope columns into the
CRD_Overview sheet of the original CRD Excel file.

Usage (called by the skill):
    python export_excel.py <results_json> <source_xlsx> <output_xlsx>

Column layout written (3 columns only):
    ┌──────────────────────────────────────────────────────────────────────┐
    │  Tool Generated Chargeability & Feasibility  (section banner row)   │
    ├─────────────────────────┬───────────────────────────────┬───────────┤
    │  Tool Assessment        │  Tool Reason                  │  Scope    │
    ├─────────────────────────┼───────────────────────────────┼───────────┤
    │  Chargeable: X |        │  <rule explanation>           │  Down-    │
    │  Feasibility: Yes |     │                               │  stream   │
    │  Complexity: Straight…  │                               │           │
    └─────────────────────────┴───────────────────────────────┴───────────┘

Styling:
  • Section banner  : black fill (#000000), white bold text, left-aligned
  • Column headers  : black fill (#000000), white bold text, left-aligned
  • Data cells      : copied fill + font from adjacent row cells (matches template)
"""

import sys
import json
import os
import openpyxl
from copy import copy
from openpyxl.styles import PatternFill, Font, Alignment, NamedStyle
from openpyxl.utils import get_column_letter

# ── Styling constants ──────────────────────────────────────────────────────────
# patternType (not fill_type) + full 8-char hex + bgColor for maximum Excel compatibility
BLACK_FILL  = PatternFill(patternType="solid", fgColor="FF000000", bgColor="FF000000")
WHITE_BOLD  = Font(bold=True, color="FFFFFFFF", name="Arial", size=10)
DATA_FONT   = Font(name="Arial", size=10)
ALIGN_LEFT  = Alignment(horizontal="left", vertical="center", wrap_text=False)
ALIGN_WRAP  = Alignment(horizontal="left", vertical="top",    wrap_text=True)

# NamedStyle for black header cells — registering in the workbook's style XML
# guarantees the fill is persisted even when row-level styles are present.
_STYLE_NAME = "CRD_BlackHeader_Tool"

# Column headers (exact 3 — no confidence columns)
HEADERS    = ["Tool Assessment", "Tool Reason", "Scope"]
COL_WIDTHS = [40, 90, 20]   # approximate character widths

# Overview sheet name (handles both underscore and space variants)
OVERVIEW_VARIANTS = ["CRD_Overview", "CRD Overview", "CRD_OVERVIEW", "CRD OVERVIEW"]


def find_overview_sheet(wb):
    for name in OVERVIEW_VARIANTS:
        if name in wb.sheetnames:
            return wb[name]
    # Fallback: first sheet containing 'overview' (case-insensitive)
    for name in wb.sheetnames:
        if "overview" in name.lower():
            return wb[name]
    raise ValueError(
        "Cannot find CRD Overview sheet. Expected one of: {}. "
        "Found: {}".format(OVERVIEW_VARIANTS, wb.sheetnames)
    )


def find_header_row(ws):
    """
    Locate the row that contains 'CRD no.' or 'CRD No.' in column A/B.
    Defaults to row 6 if not found (standard Ariba CRD template).
    """
    for row in ws.iter_rows(min_row=1, max_row=20):
        for cell in row[:3]:
            if cell.value and "crd" in str(cell.value).lower() and "no" in str(cell.value).lower():
                return cell.row
    return 6


def build_tool_assessment(r):
    return "Chargeable: {} | Feasibility: {} | Complexity: {}".format(
        r.get("chargeable", ""),
        r.get("technically_feasible", ""),
        r.get("complexity", ""),
    )


def _ensure_named_style(wb):
    """Register the black-header NamedStyle once on the workbook."""
    if _STYLE_NAME not in wb.named_styles:
        ns           = NamedStyle(name=_STYLE_NAME)
        ns.fill      = PatternFill(patternType="solid", fgColor="FF000000", bgColor="FF000000")
        ns.font      = Font(bold=True, color="FFFFFFFF", name="Arial", size=10)
        ns.alignment = Alignment(horizontal="left", vertical="center", wrap_text=False)
        wb.add_named_style(ns)


def _apply_black(cell):
    """Apply black fill + white bold font directly to a cell (belt-and-suspenders)."""
    cell.fill      = PatternFill(patternType="solid", fgColor="FF000000", bgColor="FF000000")
    cell.font      = Font(bold=True, color="FFFFFFFF", name="Arial", size=10)
    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=False)


def write_tool_assessment(ws, results, header_row):
    result_map = {r["crd_no"]: r for r in results}
    start_col  = ws.max_column + 1
    banner_row = header_row - 1   # one row above the column headers

    # Register NamedStyle on the parent workbook
    _ensure_named_style(ws.parent)

    # ── Rows 1 → (banner_row-1): extend existing black top-section background ──
    # The CRD_Overview header rows are black in the original template.
    # New columns appended after max_column don't inherit that fill, so
    # we explicitly black-fill every cell in those rows for each new column.
    for fill_row in range(1, banner_row):
        for offset in range(len(HEADERS)):
            _apply_black(ws.cell(row=fill_row, column=start_col + offset))

    # ── Banner row ────────────────────────────────────────────────────────────
    banner_cell       = ws.cell(row=banner_row, column=start_col)
    banner_cell.value = "Tool Generated Chargeability & Feasibility"
    _apply_black(banner_cell)
    for offset in range(1, len(HEADERS)):
        _apply_black(ws.cell(row=banner_row, column=start_col + offset))

    # ── Column header row ─────────────────────────────────────────────────────
    for i, header in enumerate(HEADERS):
        col = start_col + i
        c   = ws.cell(row=header_row, column=col, value=header)
        _apply_black(c)
        ws.column_dimensions[get_column_letter(col)].width = COL_WIDTHS[i]

    # ── Data rows ─────────────────────────────────────────────────────────────
    for row in ws.iter_rows(min_row=header_row + 1, max_row=ws.max_row):
        first_cell = row[0]
        if not (first_cell.value and str(first_cell.value).strip().upper().startswith("CRD")):
            continue
        crd_no = str(first_cell.value).strip()
        if crd_no not in result_map:
            continue

        r       = result_map[crd_no]
        row_num = first_cell.row

        # Copy fill and font from the first existing cell in this row so the
        # new columns match the existing row background and text style exactly.
        ref_fill = copy(first_cell.fill)
        ref_font = first_cell.font
        row_font = Font(
            bold  = False,
            size  = ref_font.size if ref_font.size else 9,
            name  = ref_font.name if ref_font.name else "Arial",
            color = ref_font.color,
        )

        c_assess = ws.cell(row=row_num, column=start_col,     value=build_tool_assessment(r))
        c_reason = ws.cell(row=row_num, column=start_col + 1, value=r.get("reason", ""))
        c_scope  = ws.cell(row=row_num, column=start_col + 2, value=r.get("scope",  ""))

        for c, align in [(c_assess, ALIGN_WRAP), (c_reason, ALIGN_WRAP), (c_scope, ALIGN_LEFT)]:
            c.fill      = ref_fill
            c.font      = row_font
            c.alignment = align


def main():
    if len(sys.argv) < 4:
        print("Usage: export_excel.py <results_json> <source_xlsx> <output_xlsx>")
        sys.exit(1)

    results_path = sys.argv[1]
    source_path  = sys.argv[2]
    output_path  = sys.argv[3]

    # Resolve output path: if only a filename is given, place in cwd
    if not os.path.isabs(output_path):
        output_path = os.path.join(os.getcwd(), output_path)

    with open(results_path, "r", encoding="utf-8") as f:
        results = json.load(f)

    wb         = openpyxl.load_workbook(source_path)
    ws         = find_overview_sheet(wb)
    header_row = find_header_row(ws)

    write_tool_assessment(ws, results, header_row)

    wb.save(output_path)
    print("Excel saved: {}".format(output_path))


if __name__ == "__main__":
    main()
