"""export_excel.py
Usage: python export_excel.py <results_json> <source_excel> <output_filename>
  results_json     Path to crd_results.json
  source_excel     Path to the original SAP Ariba CRD Excel tracker
  output_filename  Desired output filename (auto-increments if file is locked)
Embeds Tool Assessment, Tool Reason, and Scope columns into the CRD Overview sheet.
Handles all known overview sheet names: 'CRD_Overview', 'CRD Overview', 'Overview'.
Handles both old format IDs ('CRD-01') and new consolidated format IDs
('Contracts Tab - C1', 'Sourcing Tab - S2', etc.) by normalizing before lookup.
Section header and column headers use white font (inheriting source fill).
Data rows use plain black non-underlined text.
v4.8.2: Reverted to source-file fill for section/column headers; font forced WHITE
  so text is visible on any dark background without overriding the fill.
"""
import openpyxl
import openpyxl.styles
from openpyxl.styles import Font, Color
import copy
import json
import os
import re
import sys

results_path = sys.argv[1]
excel_path = sys.argv[2]
output_filename = sys.argv[3] if len(sys.argv) > 3 else 'CRD_Tool_Assessment.xlsx'

# Auto-increment output filename if locked
base, ext = os.path.splitext(output_filename)
candidate = output_filename
for i in range(2, 20):
    try:
        f = open(candidate, 'ab')
        f.close()
        break
    except PermissionError:
        candidate = f"{base}_v{i}{ext}"
output_filename = candidate

with open(results_path, 'r', encoding='utf-8') as f:
    results = json.load(f)

wb = openpyxl.load_workbook(excel_path)

# Support all known overview sheet name variants (old and new format)
overview_name = next(
    (n for n in ['CRD_Overview', 'CRD Overview', 'Overview'] if n in wb.sheetnames),
    None
)
if not overview_name:
    print('ERROR: CRD Overview sheet not found')
    sys.exit(1)
ws = wb[overview_name]

# ---------------------------------------------------------------------------
# ID normalizer: maps Overview cell values to extraction script IDs
# Old format:  'CRD-01'              -> 'CRD-01'  (unchanged)
# New format:  'Contracts Tab - C1'  -> 'C-1'
#              'Sourcing Tab - S2'   -> 'S-2'
#              'SPM Tab - SPM1'      -> 'SPM-1'
#              'Forms Tab - F1'      -> 'F-1'
#              'Savings Tab - SF1'   -> 'SF-1'
# ---------------------------------------------------------------------------
def normalize_crd_id(raw):
    if not raw:
        return None
    s = str(raw).strip()
    if ' - ' in s:
        s = s.split(' - ')[-1].strip()
    s = re.sub(r'^([A-Za-z]+)(\d+)$', r'\1-\2', s)
    return s


# ---------------------------------------------------------------------------
# Style helpers
#
# copy_style_white : copies fill + border from src, forces WHITE non-underlined font.
#                   Use for section header and column headers on dark backgrounds.
#
# copy_style_clean : copies fill + border from src, forces BLACK non-underlined font.
#                   Use for data rows on light/white backgrounds.
# ---------------------------------------------------------------------------
def copy_style_white(src, dst):
    """Copy fill, border from src but always write bold white non-underlined font.
    Designed for header rows that have a dark background fill.
    """
    if src.has_style:
        src_font = src.font
        dst.font = Font(
            name=src_font.name,
            size=src_font.size,
            bold=src_font.bold,
            italic=src_font.italic,
            underline=None,
            color=Color(rgb='FFFFFFFF'),  # always white text
        )
        dst.fill = copy.copy(src.fill)
        dst.border = copy.copy(src.border)


def copy_style_clean(src, dst):
    """Copy fill, border from src but always write plain black non-underlined font.
    Designed for data rows on light/white backgrounds.
    """
    if src.has_style:
        src_font = src.font
        dst.font = Font(
            name=src_font.name,
            size=src_font.size,
            bold=src_font.bold,
            italic=src_font.italic,
            underline=None,
            color=Color(rgb='FF000000'),  # always black text
        )
        dst.fill = copy.copy(src.fill)
        dst.border = copy.copy(src.border)


# Detect header row: look for a row containing 'CRD' and 'no' (case-insensitive)
HEADER_ROW = 6  # default
for i, row in enumerate(ws.iter_rows(max_row=15, values_only=True), 1):
    if any(row) and row[0] and 'crd' in str(row[0]).lower() and 'no' in str(row[0]).lower():
        HEADER_ROW = i
        break
DATA_START = HEADER_ROW + 1

# Find last used column in header row
last_col = 1
for cell in ws[HEADER_ROW]:
    if cell.value is not None:
        last_col = cell.column

tool_col_1 = last_col + 1  # Tool Assessment
tool_col_2 = last_col + 2  # Tool Reason
tool_col_3 = last_col + 3  # Scope

def unmerge_overlap(ws, min_c, max_c, min_r=None, max_r=None):
    to_unmerge = []
    for mr in ws.merged_cells.ranges:
        col_overlap = mr.min_col <= max_c and mr.max_col >= min_c
        row_overlap = (min_r is None) or (mr.min_row <= max_r and mr.max_row >= min_r)
        if col_overlap and row_overlap:
            to_unmerge.append(str(mr))
    for r in to_unmerge:
        ws.unmerge_cells(r)

# Unmerge anything overlapping new columns globally
unmerge_overlap(ws, tool_col_1, tool_col_3)

ref_header = ws.cell(row=HEADER_ROW, column=last_col)

# Group header row above column headers
# Inherits source fill; font is forced WHITE so it is visible on any dark background.
group_row = HEADER_ROW - 1
if group_row >= 1:
    gc = ws.cell(row=group_row, column=tool_col_1)
    gc.value = 'Tool Generated Chargeability & Feasibility'
    copy_style_white(ref_header, gc)
    try:
        ws.merge_cells(start_row=group_row, start_column=tool_col_1,
                       end_row=group_row, end_column=tool_col_3)
    except Exception:
        pass

# Column headers — white font (same dark background as section header)
for col, label in [(tool_col_1, 'Tool Assessment'), (tool_col_2, 'Tool Reason'), (tool_col_3, 'Scope')]:
    cell = ws.cell(row=HEADER_ROW, column=col)
    cell.value = label
    copy_style_white(ref_header, cell)

# Build results map keyed by CRD ID
results_map = {str(r['crd_no']): r for r in results}

# Data rows — black font on source data-row background
for row_idx in range(DATA_START, DATA_START + 300):
    crd_no_raw = ws.cell(row=row_idx, column=1).value
    if not crd_no_raw:
        continue

    normalized = normalize_crd_id(crd_no_raw)
    r = results_map.get(normalized) or results_map.get(str(crd_no_raw).strip())
    if not r:
        continue

    ref_data = ws.cell(row=row_idx, column=1)
    unmerge_overlap(ws, tool_col_1, tool_col_3, row_idx, row_idx)

    assessment = f"Chargeable: {r['chargeable']} | Feasibility: {r['technically_feasible']} | Complexity: {r['complexity']}"
    scope = r.get('scope', 'Downstream')
    for col, value in [(tool_col_1, assessment), (tool_col_2, r['reason']), (tool_col_3, scope)]:
        cell = ws.cell(row=row_idx, column=col)
        cell.value = value
        copy_style_clean(ref_data, cell)
        cell.alignment = openpyxl.styles.Alignment(wrap_text=True, vertical='top')

ws.column_dimensions[openpyxl.utils.get_column_letter(tool_col_1)].width = 40
ws.column_dimensions[openpyxl.utils.get_column_letter(tool_col_2)].width = 70
ws.column_dimensions[openpyxl.utils.get_column_letter(tool_col_3)].width = 15

wb.save(output_filename)
print(f'Saved: {output_filename}')
