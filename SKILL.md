---
name: ariba-crd-chargeability
description: >-
  Analyzes an uploaded Ariba CRD (Customization Requirement Description) Excel tracker to assess the chargeability of each custom field requirement. Supports three scopes: SAP Ariba APPS (Downstream), Ariba Network (AN), and Upstream (SAP Ariba Sourcing). For Downstream: supports Organic DD and SAP Store Pack engagement types. For AN and Upstream: single rule set each. Outputs dashboard cards in chat, a downloadable Excel file, and a styled HTML dashboard. Use this skill whenever the user uploads a CRD Excel file and asks to check chargeability, count chargeable fields, assess feasibility, or review CRD field types. Trigger phrases: "check CRD chargeability", "analyze CRD", "count chargeable fields", "CRD feasibility check", "check my CRD tracker", "Organic DD chargeability", "SAP Store Pack chargeability", "AN chargeability", "Ariba Network CRD", "check AN customizations", "upstream chargeability", "sourcing CRD", "upstream CRD".
allowed-tools: execute write_file read_file render_ui
metadata:
  author: Ariba Technology Consultant
  version: 4.9.1
  tags: ariba crd chargeability custom-fields buying-invoicing procurement feasibility complexity ariba-network an
---

# Ariba CRD Chargeability Checker

Analyze an uploaded CRD Excel tracker to assess the chargeability of each custom field requirement. Supports three deployment scopes: **SAP Ariba APPS (Downstream)**, **Ariba Network (AN)**, and **Upstream (SAP Ariba Sourcing)**. For Downstream: supports two engagement types: **Organic DD** and **SAP Store Pack**. For AN and Upstream: single rule set each. Output dashboard cards in chat, embed results in the original CRD file, and generate a styled HTML dashboard.

**The chargeable count for each CRD must be derived solely from the chargeability rules in Step 3. Do not use any pre-filled count values from the Excel file.**

---

## When to Activate

Activate when the user:
- Uploads a CRD Excel file and asks about chargeability, field count, or feasibility
- Says "check CRD chargeability", "count chargeable fields", "check my CRD", "analyze CRD chargeability"
- Asks how many CRDs are chargeable or wants a chargeability summary
- Asks whether a CRD is technically feasible or complex to implement
- Mentions AN customizations, Ariba Network CRD, or AN chargeability analysis
- Mentions Upstream CRDs, SAP Ariba Sourcing customizations, or Upstream chargeability analysis

---

## Step -1 — Bootstrap (MANDATORY — Run Before All Other Steps)

> **Version-aware cache — run this check BEFORE writing any files.**
>
> 1. `read_file` → `<scratch>/bootstrap_version.txt` (if file is missing, treat cached version as `none`)
> 2. Compare cached version against this skill's version: **`4.9.1`**
> 3. **Versions match** → skip all six `write_file` calls entirely. Proceed directly to Step 0.
> 4. **Versions differ or file missing** → write all six files (parallel turn below), then write `4.9.1` to `<scratch>/bootstrap_version.txt`.
>
> This ensures scripts and reference files are only re-written when the skill version changes — not on every run.
> Emit all six `write_file` calls in a **single parallel turn**, then wait for all to complete before writing the version file.

---

**Script 1 of 3 — `<scratch>/extract_crd.py`**

Use `write_file` to write the following content to `<scratch>/extract_crd.py`:

```python
"""extract_crd.py  (v3 - supports old format and new consolidated upstream format)
Usage: python extract_crd.py <excel_path> [crd_list]
  excel_path  Path to the SAP Ariba CRD Excel tracker
  crd_list    Comma-separated IDs (e.g. CRD-01,S-1,C-2) or 'all' (default: all)
Outputs JSON array to stdout.

Format detection:
  New format: has consolidated upstream sheets (Sourcing_Customizations etc.) + 'Overview' sheet
  Old format: individual CRD-XX sheets only (with optional 'CRD Overview' sheet)

New format upstream consolidated IDs:
  Sourcing     -> S-1, S-2  ... (Sourcing_Customizations)
  Contracts    -> C-1, C-2  ... (Contracts_Customizations)
  SPM          -> SPM-1 ...     (SPM_Customizations)
  Forms/dForms -> F-1, F-2  ... (Forms_Customizations)
  Savings      -> SF-1, SF-2 .. (Savings_Customizations)
"""
import openpyxl
import sys
import re
import json


# ---------------------------------------------------------------------------
# Adaptive field-section helpers (for CRD-XX sheets)
# ---------------------------------------------------------------------------

def find_field_section(ws):
    for row in ws.iter_rows(min_row=1, max_row=25, min_col=1, max_col=16):
        for cell in row:
            if cell.value and 'custom field detail' in str(cell.value).lower():
                h_row = cell.row
                h_col = cell.column
                for r in range(h_row + 1, h_row + 6):
                    v_same = str(ws.cell(row=r, column=h_col).value or '').strip()
                    v_next = str(ws.cell(row=r, column=h_col + 1).value or '').strip()
                    if v_same and ':' in v_same:
                        return h_row, h_col, h_col + 1
                    if v_next and ':' in v_next:
                        return h_row, h_col + 1, h_col + 2
                return h_row, h_col, h_col + 1
    return None


def extract_field_values(ws, header_row, label_col, value_col):
    label_map = {}
    for r in range(header_row + 1, header_row + 35):
        raw = str(ws.cell(row=r, column=label_col).value or '').strip()
        if not raw:
            continue
        normalized = raw.lower().rstrip(':').strip()
        label_map[normalized] = r

    def lookup(*terms):
        for term in terms:
            t = term.lower()
            for lbl, row in label_map.items():
                if lbl == t or lbl.startswith(t):
                    return ws.cell(row=row, column=value_col).value
        return None

    return {
        'document_type':        lookup('document'),
        'label':                lookup('label (key)', 'label'),
        'technical_name':       lookup('technical name'),
        'type':                 lookup('type'),
        'reportable':           lookup('reportable'),
        'erp_importable':       lookup('erp importable', 'erp import'),
        'erp_exportable':       lookup('erp exportable', 'erp export'),
        'an_field':             lookup('an field'),
        'visibility_condition': lookup('visibility condition', 'visibility'),
    }


# ---------------------------------------------------------------------------
# Format detection
# ---------------------------------------------------------------------------

NEW_FORMAT_SIGNAL_SHEETS = [
    'Sourcing_Customizations', 'Contracts_Customizations',
    'SPM_Customizations', 'Forms_Customizations', 'Savings_Customizations'
]

def detect_format(wb):
    """Returns 'new' if consolidated upstream sheets present, else 'old'."""
    for s in NEW_FORMAT_SIGNAL_SHEETS:
        if s in wb.sheetnames:
            return 'new'
    return 'old'


# ---------------------------------------------------------------------------
# Overview / status reader
# ---------------------------------------------------------------------------

def read_overview(wb, fmt):
    if fmt == 'new':
        ov_name = 'Overview' if 'Overview' in wb.sheetnames else None
    else:
        ov_name = next((n for n in ['CRD_Overview', 'CRD Overview'] if n in wb.sheetnames), None)

    overview_data = {}
    if not ov_name:
        return overview_data

    ws = wb[ov_name]
    data_start = 7
    for i, row in enumerate(ws.iter_rows(max_row=15, values_only=True), 1):
        if row[0] and 'crd' in str(row[0]).lower() and 'no' in str(row[0]).lower():
            data_start = i + 1
            break

    for row in ws.iter_rows(min_row=data_start, max_row=300, values_only=True):
        if not row[0]:
            continue
        identifier = str(row[0]).strip()
        if not identifier:
            continue
        overview_data[identifier] = {
            'title':        row[1] if len(row) > 1 else None,
            'solution':     row[2] if len(row) > 2 else None,
            'status':       row[6] if len(row) > 6 else None,
            'fields_count': row[7] if len(row) > 7 else None,
        }
    return overview_data


def get_status_for(crd_id, overview_data):
    if crd_id in overview_data:
        return overview_data[crd_id].get('status')
    crd_clean = re.sub(r'[\s\-_]', '', crd_id).lower()
    for key, val in overview_data.items():
        key_clean = re.sub(r'[\s\-_]', '', key).lower()
        if crd_clean == key_clean or key_clean.endswith(crd_clean):
            return val.get('status')
    return None


# ---------------------------------------------------------------------------
# B&I CRD-XX sheet extraction
# ---------------------------------------------------------------------------

def extract_bi_crds(wb, target_crds, overview_data):
    results = []
    for crd_name in target_crds:
        if crd_name not in wb.sheetnames:
            continue
        ws = wb[crd_name]

        crd_no         = ws['C8'].value
        title          = ws['C10'].value
        existing_field = ws['C13'].value

        section = find_field_section(ws)
        if section:
            h_row, lbl_col, val_col = section
            fv = extract_field_values(ws, h_row, lbl_col, val_col)
            doc_type   = fv['document_type']
            label      = fv['label']
            tech_name  = fv['technical_name']
            field_type = fv['type']
            reportable = fv['reportable']
            erp_import = fv['erp_importable']
            erp_export = fv['erp_exportable']
            an_field   = fv['an_field']
            vis_inline = fv['visibility_condition']
            layout_used = f'adaptive (label col {lbl_col}, value col {val_col})'
        else:
            doc_type   = ws['K10'].value
            label      = ws['K11'].value
            tech_name  = ws['K12'].value
            field_type = ws['K13'].value
            reportable = ws['K14'].value
            erp_import = ws['K15'].value
            erp_export = ws['K16'].value
            an_field   = ws['K17'].value
            vis_inline = None
            layout_used = 'fallback (K col)'

        visibility_condition = vis_inline or ws['J31'].value
        fmd_name       = ws['M10'].value
        relation_entry = ws['M20'].value

        desc_parts = []
        for r in range(15, 40):
            val = ws.cell(row=r, column=2).value
            if val and str(val).strip():
                desc_parts.append(str(val).strip())
        description = ' '.join(desc_parts)

        struck_off = False
        try:
            if (ws['C8'].font and ws['C8'].font.strikethrough) or \
               (ws['C10'].font and ws['C10'].font.strikethrough):
                struck_off = True
        except Exception:
            pass

        ov        = overview_data.get(crd_name, {})
        status    = ov.get('status') or get_status_for(crd_name, overview_data)
        cancelled = bool(status and str(status).strip().lower() in ['cancelled', 'rejected'])

        if existing_field and 'please answer' in str(existing_field).lower():
            existing_field = None

        has_field_data = bool(label or doc_type or field_type)
        fields = []
        if has_field_data:
            def clean(v):
                return str(v).strip() if v is not None else None
            fields.append({
                'document_type':    clean(doc_type),
                'label':            clean(label),
                'technical_name':   clean(tech_name),
                'type':             clean(field_type),
                'reportable':       clean(reportable),
                'erp_importable':   clean(erp_import),
                'erp_exportable':   clean(erp_export),
                'an_field':         clean(an_field),
                '_layout_detected': layout_used,
            })

        results.append({
            'crd_no':             crd_no or crd_name,
            'title':              title,
            'description':        description,
            'struck_off':         struck_off,
            'cancelled':          cancelled,
            'status':             status,
            'existing_field':     existing_field,
            'areas_impacted':     ov.get('solution', ''),
            'fields':             fields,
            'visibility_condition': str(visibility_condition).strip() if visibility_condition else None,
            'fmd':                {'name': str(fmd_name).strip()} if fmd_name else None,
            'relation_entry':     str(relation_entry).strip() if relation_entry else None,
        })
    return results


# ---------------------------------------------------------------------------
# Consolidated upstream sheet config
# ---------------------------------------------------------------------------

CONSOLIDATED_CFG = {
    'Sourcing_Customizations': {
        'scope': 'Upstream', 'document_type': 'Sourcing Header',
        'id_pattern': r'^S-\d+$',
        'header_row': 4, 'data_start': 6,
        'id_col': 2, 'title_col': 3, 'modification_col': 4,
        'label_col': 6, 'type_col': 7,
        'reportable_col': 9, 'required_col': 11,
        'condition_col': 14,
    },
    'Contracts_Customizations': {
        'scope': 'Upstream', 'document_type': 'Contract Header',
        'id_pattern': r'^C-\d+$',
        'header_row': 4, 'data_start': 6,
        'id_col': 2, 'title_col': 3, 'modification_col': 4,
        'label_col': 6, 'type_col': 7,
        'reportable_col': 9, 'required_col': 11,
        'condition_col': 14,
    },
    'SPM_Customizations': {
        'scope': 'Upstream', 'document_type': 'SPM Header',
        'id_pattern': r'^SPM-\d+$',
        'header_row': 4, 'data_start': 6,
        'id_col': 2, 'title_col': 3, 'modification_col': 4,
        'label_col': 6, 'type_col': 7,
        'reportable_col': 9, 'required_col': 11,
        'condition_col': 14,
    },
    'Forms_Customizations': {
        'scope': 'Upstream', 'document_type': 'dForm / Savings Form',
        'id_pattern': r'^F-\d+$',
        'header_row': 5, 'data_start': 7,
        'id_col': 2, 'title_col': 3, 'modification_col': 5,
        'label_col': 7, 'type_col': 9,
        'reportable_col': None, 'required_col': 11,
        'condition_col': 13,
    },
    'Savings_Customizations': {
        'scope': 'Upstream', 'document_type': 'Savings Form',
        'id_pattern': r'^SF-\d+$',
        'header_row': 2, 'data_start': 4,
        'id_col': 2, 'title_col': 3, 'modification_col': 4,
        'label_col': 7, 'type_col': 8,
        'reportable_col': None, 'required_col': 10,
        'condition_col': 11,
    },
}

PLACEHOLDER_TITLES = {
    'enter a short description', 'enter a title', '9999999', 'sample',
    'enter a title or this request. ', 'new sourcing field 1', 'new sourcing field 2',
    'new contract field 1', 'new contract field 2', 'new spm field 1', 'new spm field 2',
    'new forms field 1', 'new forms field 2', 'new savings form field 1', 'new savings form field 2',
    'new savings form field 3', 'new savings form field 4', 'new savings form field 5',
}

def is_placeholder_title(v):
    if v is None:
        return True
    s = str(v).strip()
    return s == '0' or s == '' or s.lower()[:50] in PLACEHOLDER_TITLES


def extract_consolidated_sheet(wb, sheet_name, cfg, overview_data, target_ids=None):
    """Extract customization rows from a consolidated upstream sheet."""
    if sheet_name not in wb.sheetnames:
        return []

    ws = wb[sheet_name]
    id_pat = re.compile(cfg['id_pattern'], re.IGNORECASE)
    results = []
    seen_ids = set()

    def cv(r, col):
        if col is None:
            return None
        v = ws.cell(row=r, column=col).value
        return str(v).strip() if v is not None else None

    for r in range(cfg['data_start'], ws.max_row + 1):
        raw_id = cv(r, cfg['id_col'])
        if not raw_id:
            continue
        if not id_pat.match(raw_id):
            continue
        if raw_id in seen_ids:
            continue
        seen_ids.add(raw_id)
        if target_ids and raw_id not in target_ids:
            continue

        title        = cv(r, cfg['title_col'])
        label        = cv(r, cfg['label_col'])
        field_type   = cv(r, cfg['type_col'])
        reportable   = cv(r, cfg['reportable_col']) if cfg.get('reportable_col') else None
        condition    = cv(r, cfg['condition_col'])
        modification = cv(r, cfg['modification_col'])
        required     = cv(r, cfg['required_col'])

        if is_placeholder_title(title):
            title = raw_id

        status    = get_status_for(raw_id, overview_data)
        cancelled = bool(status and str(status).strip().lower() in
                         ['cancelled but built', 'cancelled', 'rejected'])

        desc_parts = []
        if modification:
            desc_parts.append(f'Type: {modification}')
        if required:
            desc_parts.append(f'Required: {required}')
        if condition:
            desc_parts.append(f'Condition: {condition}')
        description = ' | '.join(desc_parts) if desc_parts else f'{sheet_name} customization'

        fields = []
        if label or field_type:
            fields.append({
                'document_type':    cfg['document_type'],
                'label':            label,
                'technical_name':   None,
                'type':             field_type,
                'reportable':       reportable,
                'erp_importable':   None,
                'erp_exportable':   None,
                'an_field':         None,
                '_layout_detected': f'consolidated-{sheet_name}',
            })

        results.append({
            'crd_no':             raw_id,
            'title':              title,
            'description':        description,
            'struck_off':         False,
            'cancelled':          cancelled,
            'status':             status,
            'existing_field':     'No' if modification and 'new' in str(modification).lower() else None,
            'areas_impacted':     cfg['scope'],
            'fields':             fields,
            'visibility_condition': condition,
            'fmd':                None,
            'relation_entry':     None,
        })

    return results


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

file_path = sys.argv[1]
crd_arg   = sys.argv[2].strip() if len(sys.argv) > 2 else 'all'

wb  = openpyxl.load_workbook(file_path, data_only=True)
fmt = detect_format(wb)
overview_data = read_overview(wb, fmt)

all_bi_sheets = [s for s in wb.sheetnames if re.match(r'^CRD-\d+$', s)]

results = []

if crd_arg.lower() == 'all':
    results += extract_bi_crds(wb, all_bi_sheets, overview_data)
    if fmt == 'new':
        for sheet_name, cfg in CONSOLIDATED_CFG.items():
            results += extract_consolidated_sheet(wb, sheet_name, cfg, overview_data)
else:
    requested = [c.strip() for c in crd_arg.split(',')]
    bi_req  = [r for r in requested if re.match(r'^CRD-\d+$', r, re.IGNORECASE)]
    ups_req = [r for r in requested if not re.match(r'^CRD-\d+$', r, re.IGNORECASE)]

    results += extract_bi_crds(wb, bi_req, overview_data)

    if fmt == 'new' and ups_req:
        for sheet_name, cfg in CONSOLIDATED_CFG.items():
            results += extract_consolidated_sheet(wb, sheet_name, cfg, overview_data,
                                                  target_ids=set(ups_req))

print(json.dumps(results, indent=2, default=str))
```

---

**Script 2 of 3 — `<scratch>/export_excel.py`**

Use `write_file` to write the following content to `<scratch>/export_excel.py`:

```python
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
# ---------------------------------------------------------------------------
def copy_style_white(src, dst):
    """Copy fill, border from src but always write bold white non-underlined font."""
    if src.has_style:
        src_font = src.font
        dst.font = Font(
            name=src_font.name,
            size=src_font.size,
            bold=src_font.bold,
            italic=src_font.italic,
            underline=None,
            color=Color(rgb='FFFFFFFF'),
        )
        dst.fill = copy.copy(src.fill)
        dst.border = copy.copy(src.border)


def copy_style_clean(src, dst):
    """Copy fill, border from src but always write plain black non-underlined font."""
    if src.has_style:
        src_font = src.font
        dst.font = Font(
            name=src_font.name,
            size=src_font.size,
            bold=src_font.bold,
            italic=src_font.italic,
            underline=None,
            color=Color(rgb='FF000000'),
        )
        dst.fill = copy.copy(src.fill)
        dst.border = copy.copy(src.border)


# Detect header row
HEADER_ROW = 6
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

unmerge_overlap(ws, tool_col_1, tool_col_3)

ref_header = ws.cell(row=HEADER_ROW, column=last_col)

# Group header row
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

# Column headers
for col, label in [(tool_col_1, 'Tool Assessment'), (tool_col_2, 'Tool Reason'), (tool_col_3, 'Scope')]:
    cell = ws.cell(row=HEADER_ROW, column=col)
    cell.value = label
    copy_style_white(ref_header, cell)

# Build results map keyed by CRD ID
results_map = {str(r['crd_no']): r for r in results}

# Data rows
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
```

---

**Script 3 of 3 — `<scratch>/generate_html.py`**

Use `write_file` to write the following content to `<scratch>/generate_html.py`:

```python
"""generate_html.py
Usage: python generate_html.py <results_json> <output_html>
  results_json  Path to crd_results.json
  output_html   Output HTML file path (saved to working directory)
Generates a styled HTML chargeability dashboard.
"""
import json
import sys

results_path = sys.argv[1]
output_path = sys.argv[2]

with open(results_path, 'r', encoding='utf-8') as f:
    results = json.load(f)

total = len(results)
total_chargeable = sum(r['chargeable'] for r in results)
avg_c = round(sum(r['chargeability_confidence'] for r in results) / total)
avg_f = round(sum(r['feasibility_confidence'] for r in results) / total)
avg_conf = round((avg_c + avg_f) / 2)
feas_yes = sum(1 for r in results if r['technically_feasible'] == 'Yes')
feas_no = sum(1 for r in results if r['technically_feasible'] == 'No')
feas_rev = sum(1 for r in results if r['technically_feasible'] == 'Needs Review')
cmpl_s = sum(1 for r in results if r['complexity'] == 'Straightforward')
cmpl_m = sum(1 for r in results if r['complexity'] == 'Medium')
cmpl_c = sum(1 for r in results if r['complexity'] == 'Complex')

ds_count = sum(1 for r in results if r.get('scope', 'Downstream') == 'Downstream')
an_count = sum(1 for r in results if r.get('scope', 'Downstream') == 'AN')
up_count = sum(1 for r in results if r.get('scope', 'Downstream') == 'Upstream')
scope_types = [s for s, c in [('Downstream', ds_count), ('AN', an_count), ('Upstream', up_count)] if c > 0]
has_multiple = len(scope_types) > 1

def scope_display(s):
    return {'Downstream': 'SAP Ariba B&amp;I', 'AN': 'Ariba Network (AN)', 'Upstream': 'Upstream (Sourcing)'}.get(s, s)

if has_multiple:
    scope_label = ' &amp; '.join(scope_display(s) for s in scope_types)
    scope_footer = ' / '.join(scope_types)
elif an_count > 0:
    scope_label = 'Ariba Network (AN)'
    scope_footer = 'AN'
elif up_count > 0:
    scope_label = 'Upstream (SAP Ariba Sourcing)'
    scope_footer = 'Upstream'
else:
    scope_label = 'SAP Ariba B&amp;I (Downstream)'
    scope_footer = 'Downstream'

engagement = ''
for r in results:
    if r.get('scope', 'Downstream') == 'Downstream' and r.get('engagement_type') and r['engagement_type'] not in ('AN', 'Upstream'):
        engagement = r['engagement_type']
        break

first_crd = results[0]['crd_no'] if results else ''
last_crd = results[-1]['crd_no'] if results else ''
crd_range = f"{first_crd} to {last_crd}" if first_crd != last_crd else first_crd

def conf_color(v):
    if v >= 90: return '#22c55e'
    if v >= 70: return '#f59e0b'
    if v >= 50: return '#f97316'
    return '#ef4444'

def badge(val, colors):
    c = colors.get(val, '#94a3b8')
    return f'<span style="background:{c};color:#fff;padding:2px 10px;border-radius:12px;font-size:12px;font-weight:600;">{val}</span>'

FEAS_COLORS = {'Yes': '#22c55e', 'No': '#ef4444', 'Needs Review': '#f59e0b'}
CMPL_COLORS = {'Straightforward': '#22c55e', 'Medium': '#f59e0b', 'Complex': '#ef4444'}
SCOPE_COLORS = {'Downstream': '#1e40af', 'AN': '#7c3aed', 'Upstream': '#0d9488'}

scope_kpi = ''
if has_multiple:
    parts = []
    if ds_count > 0: parts.append(f'{ds_count} DS')
    if an_count > 0: parts.append(f'{an_count} AN')
    if up_count > 0: parts.append(f'{up_count} UP')
    scope_kpi = f'<div class="kpi"><div class="val">{parts[0]}</div><div class="lbl">Scope Split</div><div class="sub">{"&middot;".join(parts[1:])}</div></div>'

eng_sub = f'Engagement: {engagement} &nbsp;|&nbsp; ' if engagement else ''

rows_html = ''
for r in results:
    cc, fc = r['chargeability_confidence'], r['feasibility_confidence']
    reason_text = r['reason'][:200] + ('...' if len(r['reason']) > 200 else '')
    scope_val = r.get('scope', 'Downstream')
    scope_color = SCOPE_COLORS.get(scope_val, '#94a3b8')
    scope_badge = f'<span style="background:{scope_color};color:#fff;padding:2px 8px;border-radius:8px;font-size:12px;font-weight:600;">{scope_val}</span>'
    rows_html += f"""
    <tr>
      <td style="font-weight:700;color:#1e40af;">{r['crd_no']}</td>
      <td>{r['title']}</td>
      <td style="text-align:center;">{scope_badge}</td>
      <td><span style="background:#e0e7ff;color:#3730a3;padding:2px 8px;border-radius:8px;font-size:12px;">{r['field_type']}</span></td>
      <td style="text-align:center;font-weight:700;font-size:16px;">{r['chargeable']}</td>
      <td style="text-align:center;">{badge(r['technically_feasible'], FEAS_COLORS)}</td>
      <td style="text-align:center;">{badge(r['complexity'], CMPL_COLORS)}</td>
      <td style="font-size:12px;color:#475569;">{reason_text}</td>
      <td style="text-align:center;">
        <div style="font-size:12px;font-weight:600;color:{conf_color(cc)};">C: {cc}%</div>
        <div style="font-size:12px;font-weight:600;color:{conf_color(fc)};">F: {fc}%</div>
      </td>
    </tr>"""

html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>CRD Chargeability Report</title>
<style>
  * {{ box-sizing:border-box;margin:0;padding:0; }}
  body {{ font-family:'Segoe UI',Arial,sans-serif;background:#f1f5f9;color:#1e293b; }}
  .header {{ background:linear-gradient(135deg,#1e40af 0%,#3b82f6 100%);color:white;padding:32px 40px; }}
  .header h1 {{ font-size:26px;font-weight:700; }}
  .header p {{ font-size:14px;opacity:.85;margin-top:4px; }}
  .kpi-grid {{ display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:16px;padding:24px 40px; }}
  .kpi {{ background:white;border-radius:12px;padding:20px;box-shadow:0 1px 4px rgba(0,0,0,.07);text-align:center; }}
  .kpi .val {{ font-size:32px;font-weight:800;color:#1e40af; }}
  .kpi .lbl {{ font-size:13px;color:#64748b;margin-top:4px;font-weight:500; }}
  .kpi .sub {{ font-size:11px;color:#94a3b8;margin-top:2px; }}
  .table-wrap {{ padding:0 40px 40px; }}
  .table-wrap h2 {{ font-size:18px;font-weight:700;margin-bottom:16px; }}
  table {{ width:100%;border-collapse:collapse;background:white;border-radius:12px;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,.07); }}
  th {{ background:#1e40af;color:white;padding:12px 14px;font-size:12px;text-align:left;font-weight:600; }}
  td {{ padding:12px 14px;font-size:13px;border-bottom:1px solid #e2e8f0;vertical-align:top; }}
  tr:last-child td {{ border-bottom:none; }}
  tr:nth-child(even) td {{ background:#f8fafc; }}
  .footer {{ padding:20px 40px;color:#94a3b8;font-size:12px;text-align:center; }}
</style>
</head>
<body>
<div class="header">
  <h1>CRD Chargeability Report</h1>
  <p>{eng_sub}Scope: {scope_label} &nbsp;|&nbsp; {total} CRDs evaluated ({crd_range})</p>
</div>
<div class="kpi-grid">
  <div class="kpi"><div class="val">{total}</div><div class="lbl">CRDs Evaluated</div><div class="sub">{crd_range}</div></div>
  <div class="kpi"><div class="val">{total_chargeable}</div><div class="lbl">Total Chargeable</div><div class="sub">field units</div></div>
  <div class="kpi"><div class="val">{avg_conf}%</div><div class="lbl">Avg. Confidence</div><div class="sub">chargeability &amp; feasibility</div></div>
  <div class="kpi"><div class="val">{feas_yes} Yes</div><div class="lbl">Feasibility</div><div class="sub">{feas_no} No &middot; {feas_rev} Needs Review</div></div>
  <div class="kpi"><div class="val">{cmpl_m} Med</div><div class="lbl">Complexity</div><div class="sub">{cmpl_s} Strt &middot; {cmpl_c} Complex</div></div>
  {scope_kpi}
</div>
<div class="table-wrap">
  <h2>Chargeability &amp; Feasibility Results</h2>
  <table>
    <thead><tr>
      <th>CRD No.</th><th>Title</th><th>Scope</th><th>Field Type</th>
      <th>Chargeable</th><th>Feasible</th><th>Complexity</th>
      <th>Reason</th><th>Confidence</th>
    </tr></thead>
    <tbody>{rows_html}</tbody>
  </table>
</div>
<div class="footer">Generated by Joule CRD Chargeability Checker &middot; {scope_footer} &middot; {total} CRDs</div>
</body></html>"""

with open(output_path, 'w', encoding='utf-8') as f:
    f.write(html)
print(f'HTML saved: {output_path}')
```

---

**Reference 1 of 3 — `<scratch>/an-chargeability-rules.md`**

Use `write_file` to write the following content to `<scratch>/an-chargeability-rules.md`:

# SAP Ariba Chargeability Rules — Ariba Network (AN) UI Customizations

Source: MW_AribaNetwork guidelines (Procurement - Chargeable Custom Fields Delivery Guidelines)

**Scope: Ariba Network (AN) UI Customizations only — single rule set, no engagement type distinction.**

> Note: These guidelines are for TL use when building AN customizations as part of B&I deployment. Customizations on OrderConfirmation and ASN do not fall under DD scope.

---

## AN Chargeability Rules

| Field Type | Field Characteristics | Chargeable Count | Notes |
|---|---|:---:|---|
| New Fields | Add Extrinsic | **1** | 1 per Document Type. Not applicable to PO UI (read-only). The cXML PO received by the SBN must contain any needed Extrinsics — all Extrinsics on the cXML PO are displayed by default and typically do not require UI customizations. |
| New Fields | B&I AN Export Fields | **0** | New Extrinsic fields for integration — 0 count if the exported custom field on B&I is already counted as a customized field. |
| Logic Fields | New field | **0** | 0 count for a new field. If configured at same time as Extrinsic is Added, included as part of the Add Extrinsic process — not an extra chargeable count. |
| Logic Fields | Ease of Access | **1** | Fields introduced to easily compute or access in header level expressions or similar scenarios. |
| New Fields | Visibility, Validity, Editability on new fields | **0** | V/V/E defined on new fields = 0. |
| Previously Customized Fields | Customize a Previously Customized Field | **1** | If a field has previously been customized and the customer now needs to modify existing customizations — adding, removing, or changing any customizations OR adding a new additional customization = 1. |
| Invoice UI | Hiding Invoice UI Sections | **1** | Per section hidden (e.g., Tax, Shipping). Always hides the entire section — no dynamic hiding supported. |
| Invoice UI | Resequencing Invoice UI Sections | **1** | Changing the display sequence of sections on the Invoice UI. When a section is moved, all fields within it move automatically. |
| Columnar Table Display | Customize PO Party Info Table | **1** | Customizing Header text, detail column content, or hiding the ShipTo/BillTo/RemitTo party table column. |
| Columnar Table Display | Move/Add/Remove Columns or customize column headings | **1** | Applies to Line Items sections of PO/INV/SES. |
| Display Objects | Warning Messages/Notes (plain text) | **1** | Only display on UI screens and PDFs. No rules support — only static Hidden property. |
| Display Objects | Notes displayed as HTML | **2** | 1 for adding the Note object + 1 for creating the HTML content. |
| OOTB Fields | Visibility, Validity, Editability | **1** | Any V/V/E customization on OOTB fields = 1. |
| OOTB Fields | Additional Properties | **1** | Only 0 if also making other customizations to the same field. |
| OOTB Fields | Choice values (dropdown) | **1** | Setting OOTB field as a dropdown menu = 1. |
| OOTB Fields | Hide field or section | **1** | Only 0 if also making other customizations to the same field. |
| Invoice Sub-Type | Invoice customizations based on Invoice Sub-Type | **0** | Applying customizations to specific sub-types vs all does not count as additional customizations. |
| PO UI | Customize PO Header Customer Address Info | **1** | Customizing display of address information in the PO Header Customer Address section. |
| PO UI | Hide the Purchase Order section on PO UI | **1** | Only supported customization for this section is hiding it entirely. |
| PO UI | Hide PO Header Supplier Address Info | **1** | Supplier address info can be hidden from PO UI Header. |
| Invoice Drop-Down | Hide menu items in Add to Header drop-down | **1** | Removing OOTB menu items such as Shipping or Special handling. |
| Invoice Drop-Down | Hide menu items in Line Item Actions drop-down | **2** | Two separate instances (regular and Edit mode) must both be customized. |
| Invoice UI | Hide the Line Item Options mass pre-fill section | **1** | Must also hide if hiding same items from Line Item Actions drop-down. |
| Invoice PDFs | Hide Not A Tax Invoice verbiage for Australia | **1** | One of 4 PDF versions. |
| Invoice PDFs | Hide Not A Tax Invoice verbiage for South Africa | **1** | One of 4 PDF versions. |
| Invoice PDFs | Hide Not A Tax Invoice verbiage for Colombia | **1** | One of 4 PDF versions. |
| Invoice PDFs | Hide Not A Tax Invoice verbiage for all other countries | **1** | One of 4 PDF versions. |
| Other | Invoice ID Length (99.9% of time) | **0** | Use Document Number Preferences feature — no UI customizations required in 99.9% of cases. |
| Other | Invoice ID Validation (.01% exception) | **1** | Only when DNP feature cannot support the requirement. |
| Other | Order Confirmation ID Length Limitation | **1** | Common request. |
| Other | Legacy PO Verbiage Customization | **1** | This purchase order has already been fulfilled verbiage. |
| Other | PO Footer | **1** | Customize what displays in the PO Footer. |
| Adding Tariff Charges to Invoice | Add Invoice Header Level Tariff Charges (UI Custs) | **1** | Only UI customization needed is to remove "Allowance" from the "Add to Header" drop-down. |
| Adding Tariff Charges to Invoice | Add Invoice LI Tariff Charges (UI Custs) | **1** | Only UI customization needed is to remove "Allowance" from the "Line Item Actions" drop-down. |
| Adding Tariff Charges to Invoice | Add Invoice LI Tariff Charges (Customer) | **0** | Customer task only — not a UI customization. |
| Flipping Service PO LI Tax to Invoice | UI Customizations | **1** | Special Invoice UI customization required to flip LI tax from Parent lines into Invoice child lines. |
| Flipping Service PO LI Tax to Invoice | Customer Configurations/Tasks | **0** | Customer task only — not a UI customization. |
| UI Layouts | New UI Layout | **1** | One count per new UI Layout created. |
| UI Layouts | Customizations deployed on a new layout | **(per rule)** | Each customization on a new layout is independently chargeable under its own applicable rule — including customizations replicated from an existing layout. |
| Supplier Groups | New Supplier Group | **1** | Creating a new supplier group to associate a UI Layout with a specific subset of suppliers = 1. |

---

## Not Customizable — AN UI (Mark as Not Feasible)

CRDs falling under any of the following categories must be marked **Not Feasible** for AN scope:

- **Font and text formatting** — font-weight, font-size, carriage returns or line feeds on any text fields
- **Invoice Header Level Credit Memo UI** — controlled by code stream, not customizable
- **Integrated OC/ASN/SES/INV documents** — cXML, EDI, ICS, and CSV uploads do not use Input UI screens
- **Localizations** — displaying static text in preferred browser language; not supported in Customization Packs via SAP Store
- **Buyer Logo** — Account Level configuration, not a UI customization
- **Search Results UI screens** — PO and Invoice search results including the Actions column
- **Company Profile data access** — cannot access Buyer or Supplier Company Profile information via UI customizations
- **Buyer/Supplier Account configurations** — UI customizations can only access the cXML PO and field values on the Input UI
- **Non-PO Invoice Ship To drop-down** — auto-populated from Company Profile; address selection drop-down not supported
- **Clickable Buttons** — except Add to Header and Line Item Actions on the Invoice Input UI
- **Standard PDF direct customizations** — PDFs are indirectly customized when the related UI screen is customized
- **Auto-generated email notifications** — emails are not CommAuto document UI screens
- **SBN Reports** — all reports have pre-designed formats and cannot be customized
- **Custom Mappings in Managed Gateway for Spend and Network** — not included in UI Customizations
- **Multi-columnar table format for multiple sets of Extrinsic values** — not supported
- **Supplier-only UI customizations** — AN UI customizations are static and apply identically to all suppliers. Mark as Not Feasible. Note: a two-layout + supplier group approach may be an alternative — reassess using UI Layout and Supplier Group rules.
- **OC Line Item Non-Extrinsic field customizations** — only Extrinsic customizations supported on OC Line Item UI
- **ASN UI sections with known limitations** — certain sections do not support customizations

---

## Multi-Site Notes

AN chargeability is assessed per CRD. Multi-site rules follow the same principles as Downstream — apply site multipliers manually after getting per-CRD counts.

---

**Reference 2 of 3 — `<scratch>/gb-customization-limitations.md`**

Use `write_file` to write the following content to `<scratch>/gb-customization-limitations.md`:


# SAP Ariba Guided Buying — Customization Limitations

Source: Ariba Wiki – Guided Buying UTP Flows (3) > Customizations section  
Document Date: September 2026

## Limitation Types
- **Not Supported** – Completely unsupported in Guided Buying.
- **Hard-coded** – Fixed by the GB UI/code; cannot be altered via AML or configuration.
- **Limited** – Partial support — available only in specific contexts (e.g. checkout page only).

---

## 1. Fields on the Non-catalog Request Page

| Sub-area / Field | Limitation Description | Type | Workaround |
|---|---|---|---|
| Hard-coded standard fields | Standard fields (Product name, Service name, Category, Description, Start/End date, Quantity, UOM, Price/Currency, Expected/Max/Service amount, Supplier) are hard-coded in the GB UI and are not customizable. | Hard-coded | Design a line-item form (ReqForm) with form widgets that map to standard ReqLineItem fields. |
| Label & UI customization | No label changes, no help tips, no field requirement changes, no custom validity/editability/visibility conditions, no hiding fields, no reordering of fields. | Not Supported | Use a ReqForm (line-item form) instead of AML customizations. |
| Relation entries | No relation entries for Category, Unit of Measure, or Supplier fields (e.g. CHR-5339). | Not Supported | None documented. |
| Quantity decimal precision | Quantity field supports only up to 2 decimal places (GB-12193). | Hard-coded | None – product limitation. |
| Thousands/decimal separators | Separators are hard-coded by locale (locale.service.ts). Cannot be changed. | Hard-coded | None – inform customers of locale-driven behavior. |
| Commodity code lists | No support for separate lists of commodity codes based on Goods or Services selection. | Not Supported | Refer to configurations supported by guided buying parameters. |
| Currency dropdown | No option to hide or deactivate currencies from the currency dropdown on the Price field. | Not Supported | See 'Can I deactivate specific currencies in guided buying?' for a potential workaround. |
| Custom fields | Custom fields cannot be added to the Non-catalog request page. | Not Supported | Design a ReqForm to capture custom data instead. |

---

## 2. Fields on the Catalog Item Details Page

| Sub-area / Field | Limitation Description | Type | Workaround |
|---|---|---|---|
| Field visibility & ordering | The page is hard-coded; you cannot change which fields are visible or their display order. | Hard-coded | None – standard page only. |
| Top section layout | Top section (catalog image + product info) cannot be modified. | Hard-coded | None. |
| Product information section | The Product information section is fixed and cannot be reordered or customized. | Hard-coded | None. |

---

## 3. Supplier / Vendor Chooser

| Sub-area / Field | Limitation Description | Type | Workaround |
|---|---|---|---|
| Label changes | No label changes — except when the field appears in the line-item section of the requisition checkout page. | Limited | Label changes supported only on the requisition checkout page line-item section. |
| Field requirement changes | No field requirement changes supported. | Not Supported | None. |
| Validity / editability / visibility conditions | Limited — conditions supported only on the requisition checkout page, not on the Non-catalog request page or forms. | Limited | Implement validation policies before the add-to-cart process. |
| Section position on forms | No moving the Supplier section on forms. | Not Supported | None. |
| Relation entries | No relation entries supported. | Not Supported | None. |
| Chooser fields & field order | No changes to the fields or field order in the chooser. | Not Supported | None — GB supplier search runs through Supplier Management (SM). |
| Supplier section on externally managed forms | Cannot remove the Supplier section on externally managed forms. | Not Supported | None. |

---

## 4. Category / Commodity Code Chooser

| Sub-area / Field | Limitation Description | Type | Workaround |
|---|---|---|---|
| Label changes | Limited label changes — supported only on the requisition checkout page. | Limited | None outside checkout page. |
| Field requirement changes | No field requirement changes supported. | Not Supported | None. |
| Validity / editability / visibility conditions | Limited — supported only on the requisition checkout page, not on the Non-catalog request page or forms. | Limited | None outside checkout page. |
| Relation entries | No relation entries supported. | Not Supported | None. |
| Chooser fields & field order | No changes to the fields or field order in the chooser, except configurations supported by guided buying parameters. | Limited | Check guided buying parameters for available configurations. |
| Goods vs Services commodity code lists | No support for separate commodity code lists based on Goods or Services selection on the Non-catalog request page. | Not Supported | None — GB commodity search runs through MDS on HANA. |

---

## 5. Ad Hoc / Personal Shipping Addresses

| Sub-area / Field | Limitation Description | Type | Workaround |
|---|---|---|---|
| Complex requirements & validations | Complex requirements and validations related to ad hoc address creation are discouraged. | Limited | Customers with unsupported requirements should submit an improvement request to SAP Product Management. |

---

## 6. Field Choosers (General)

| Sub-area / Field | Limitation Description | Type | Workaround |
|---|---|---|---|
| UniqueName (ID) column removal | Cannot remove the UniqueName (ID) column from flex master data (FMD) choosers. | Not Supported | None. |
| Custom chooser columns | Limited support for defining custom columns for choosers (ChooserGroupField and HasCustomChooserGroup templates in AML file). | Limited | Test on a case-by-case basis; limited AML template support available. |
| Chooser column order | No support for defining a set order for chooser columns (predecessor property in AML file). | Not Supported | None. |

---

## 7. Fields on the Requisition Checkout Page

| Sub-area / Field | Limitation Description | Type | Workaround |
|---|---|---|---|
| Line item summary — hard-coded fields | Fields in the line item summary section are not customizable (ShortName, Quantity, UOM, Price, Net Amount, Gross Amount). | Hard-coded | None. |
| Description field label | The Description field label (LineItemProductDescription class) is non-editable. | Not Supported | None. |
| Need By Date defaulting expressions | Defaulting expressions for Need By Date are not supported in guided buying (CHR-2290, CHR-2433, CHR-2233, CHR-2701). | Limited | Making the field required is supported; avoid defaulting expressions. |
| Multiline text field character restriction | Cannot restrict multiline text fields to fewer than 1000 characters via the character width property. | Not Supported | None — rely on back-end validation if needed. |
| Tool tips | Tool tips might not display for fields if the user does not have edit access to the field. | Limited | Ensure field edit access for users who need to see tool tips. |
| Boolean fields as checkbox | Configuring a Boolean field to appear as a checkbox or to show customized Yes/No labels is not supported in guided buying. | Not Supported | Works as expected in SAP Ariba Buying; not supported in guided buying. |

---

## 8. Purchase Order Display

| Sub-area / Field | Limitation Description | Type | Workaround |
|---|---|---|---|
| PO view customization | Customers cannot customize the purchase order view in guided buying. No custom fields are shown. | Not Supported | Users can click 'View on SAP Ariba Procurement' to see the customized layout in SAP Ariba Buying. |

---

## 9. Receipt Customizations

| Sub-area / Field | Limitation Description | Type | Workaround |
|---|---|---|---|
| Custom fields & conditions on receipts | Custom fields or validity/editability/visibility conditions on Receipt, ReceiptItem, or ReceivableLineItem classes are not supported in guided buying. | Not Supported | Enable the SET_ADVANCED_RECEIVE_TAB parameter to redirect users to SAP Ariba Buying for receiving. |

---

## 10. Your Requests / Your Approvals

| Sub-area / Field | Limitation Description | Type | Workaround |
|---|---|---|---|
| Search filters & result columns | No search filter or search result column additions in the subsections of Your requests and Your approvals. | Not Supported | Use procurement workspace project field configuration for limited column control. |
| Field label changes | No field label changes supported. | Not Supported | None. |
| Look-and-feel changes | No changes to the look-and-feel of the subsections. | Not Supported | None. |
| Advanced requisition search | Advanced search options for requisitions are enabled only for customer J&J. | Limited | See Hidden parameters documentation for available workarounds. |
| Receiving page (To Receive section) | The receiving page in guided buying is not customizable at all. | Not Supported | Enable the SET_ADVANCED_RECEIVE_TAB guided buying parameter to redirect users to SAP Ariba Buying. |

---

## 11. Improved Search Results Page

| Sub-area / Field | Limitation Description | Type | Workaround |
|---|---|---|---|
| Search results tabs | It is not possible to remove, rename, or reorganize search results tabs. | Not Supported | None. |

---

## 12. "Guided Buying" Header Bar Text

| Sub-area / Field | Limitation Description | Type | Workaround |
|---|---|---|---|
| Default header text | There is no option to hide or change the default 'Guided Buying' text in the guided buying header bar. | Hard-coded | None. |

---

## 13. User Preferences Menu

| Sub-area / Field | Limitation Description | Type | Workaround |
|---|---|---|---|
| App Settings menu options | Menu options under App Settings are visible to all guided buying users. Cannot hide these options or make them non-editable. | Not Supported | None. |

---

## 14. RFQs

| Sub-area / Field | Limitation Description | Type | Workaround |
|---|---|---|---|
| Respond by date default | The Respond by date on RFQ form tiles is always defaulted to 7 days. Hard-coded with no option to customize. | Hard-coded | None documented. |
| RFQ details page / quote details layout | No options to customize the layout of fields when viewing a submitted RFQ. Displayed fields for incoming supplier quotes are fixed. | Not Supported | More customization options planned with the switch to API architecture. |


---

**Reference 3 of 3 — `<scratch>/upstream-chargeability-rules.md`**

Use `write_file` to write the following content to `<scratch>/upstream-chargeability-rules.md`:

# SAP Ariba Chargeability Rules — Upstream (SAP Ariba Sourcing)

Source: Upstream Feasibility Rules guidelines.

**Scope: Upstream (SAP Ariba Sourcing) only — single rule set, no engagement type distinction.**

> Note: 95% of the time, the number of fields = the number of customizations. One field = one customization regardless of how many changes are made to it. If two customizations are needed, it is typically because a workaround is required to fulfil the feasibility of the request.

---

## What to Check Before Scoring

Review the CRD description to identify which area(s) are impacted:

- Edits to Standard fields tab?
- Edits to Sourcing fields tab?
- Edits to Contract fields tab?
- Edits to SLP tab?
- Edits to Savings Form fields?
- Edits to dForm fields?
- New Field or Modification?
- Number of new fields / modifications?

---

## Upstream Chargeability Rules

### New Fields

| Field Type | Characteristic | Chargeable Count | Notes |
|---|---|:---:|---|
| New Fields | Create new field | **1** | New field with all associated settings — including any Visibility, Validity, or Editability conditions set at creation time. Conditions configured when building a new field are part of the field creation and are NOT counted as a separate customization. Translations for new fields are also included and not counted separately. |
| New Fields | Create new field to replace existing field | **2** | 1 customization for building the new field + 1 customization to hide the existing field. |
| New Fields | Translations for new fields | **0** | Included as part of the new field request — not a separate customization. |
| New Fields | V/V/E condition set at new field creation time | **0** | Visibility, Validity, or Editability conditions configured when creating a new field are included in the Create new field count (1 total). Do NOT apply the Field Modifications New condition rule to new fields. |

### Field Modifications (Custom or OOTB)

| Field Type | Characteristic | Chargeable Count | Notes |
|---|---|:---:|---|
| Field Modifications | Make Field Required | **1** | Applies to both custom and OOTB fields. |
| Field Modifications | Update Label or Help Tip Text | **1** | Applies to both custom and OOTB fields. |
| Field Modifications | Label Change | **1** | Applies to both custom and OOTB fields. |
| Field Modifications | Hide Field | **1** | Applies to both custom and OOTB fields. |
| Field Modifications | Edits to Advanced Settings | **1** | Any edits to advanced settings on a field = 1 customization. |
| Field Modifications | New condition (Visibility, Editability, Validity) on EXISTING field | **1** | Adding a new V/V/E condition to an already-existing field = 1 customization. This rule applies ONLY to existing fields — NOT to new fields where the condition is set at creation time. |
| Field Modifications | Sync existing field with existing fields in another form (custom dForm or Savings Form) | **1** | If a field exists in both the workspace and the form, advanced settings need to be adjusted to sync/carry over. |
| Field Modifications | Translations for existing fields | **1** | Translation to an existing live field = 1 customization. Note: multiple translations sometimes = 2 translations per 1 customization — assess case by case. |

---

## Key Principles

- **One field = one customization** regardless of the number of individual changes made to that field.
- **New field with V/V/E condition = 1**, not 2 — the condition is part of building the field.
- **Adding V/V/E to an existing field = 1** — the modification itself counts as 1 customization.
- If two customizations are needed, it is typically because a workaround is required to fulfil the feasibility of the request.

---

## Feasibility Notes

- Expressions and conditions are sometimes not feasible — highly situational.
- When feasibility of an expression or condition cannot be confirmed, flag as **Needs Review**.
- Customers often fill out CRD workbooks incorrectly (wrong tab, all customizations in one line) — flag ambiguous entries in the Reason field.

---

## Multi-Site Notes

Upstream chargeability is assessed per CRD. Apply site multipliers manually after getting per-CRD counts.

---

After all six `write_file` calls complete successfully, proceed to **Step 0**.

---

## Step 0 - Select Scope and Engagement Type

> **FAST-PATH ENTRY:** If the user directly names a specific CRD (e.g. "check CRD-01 chargeability") AND the scope and engagement type were already confirmed earlier in this conversation, skip Step 0 prompts and Step 1b selection entirely. Extract only that named CRD, apply the rules, and proceed directly to Step 3 → Step 4 output. Do NOT ask questions that were already answered.

**Step 0a - Select scope:**

Ask:
> **Which scope are you analyzing?**
> 1. Downstream (SAP Ariba APPS)
> 2. Ariba Network (AN)
> 3. Upstream (SAP Ariba Sourcing)
> 4. Multiple (combination of the above)

Wait for response. Store as `scope` - `"Downstream"`, `"AN"`, `"Upstream"`, or `"Multiple"`.

**Step 0b - Select engagement type (Downstream only):**

If `scope` is `"Downstream"` or `"Multiple"`, ask:
> **Which engagement type applies to the Downstream CRDs?**
> 1. Organic DD
> 2. SAP Store Pack

Wait for response. Store as `engagement_type`.

If `scope` is `"AN"` or `"Upstream"` only, skip this question.

---

## Step 1 - Confirm the File

Check that the user has uploaded a CRD Excel tracker file. Expected structure:
- A **CRD Overview** sheet listing all CRD numbers, titles, areas impacted, Status, and current chargeable values
- Individual **CRD-XX** sheets (e.g., CRD-01, CRD-02) with detailed field specs

If no file is available, ask the user to upload the CRD Excel tracker file before proceeding.

---

## Step 1b - Extract All CRDs and Ask Which to Evaluate

> **SINGLE-PASS: Extract and discover in one step - do NOT run a separate discovery scan first.**

Run the extraction script for **all** CRDs in one pass:

```
execute -> pip install openpyxl
execute -> python "<scratch>/extract_crd.py" "<path_to_uploaded_file>" "all"
```

Store the full JSON output as `all_extracted`. Then:

1. Filter `all_extracted` to build the **available list** - keep only CRDs where ALL of these are true:
   - `struck_off: false`
   - `cancelled: false`
   - `fields` array is not empty

2. Present the available CRDs to the user:

> **Available CRDs:**
> - CRD-01 - Material Code
> - CRD-02 - Material Usage
>
> Which would you like to evaluate? List specific ones (e.g., "CRD-01, CRD-03") or type **"all"**.

3. **Wait for the user's response.**

- **"all"** -> use the full available list from `all_extracted`
- Specific list -> filter `all_extracted` to only those CRDs
- If a named CRD is not in the available list, flag it and ask for confirmation

**Step 1c - Assign scope to CRDs (only when scope is "Multiple"):**

If `scope` is `"Multiple"`, after the user selects which CRDs to evaluate, ask:
> **Which of these CRDs are Ariba Network (AN) customizations?**
> List specific ones (e.g., "CRD-03, CRD-05") or say **"none"**.

Wait for response. Then ask:
> **Which of these CRDs are Upstream (SAP Ariba Sourcing) customizations?**
> List specific ones (e.g., "CRD-06, CRD-07") or say **"none"**.

All remaining selected CRDs not listed as AN or Upstream will be treated as Downstream.

Build `crd_scope_map` - a lookup from CRD number to scope.

If `scope` is `"Downstream"` -> all selected CRDs are Downstream.
If `scope` is `"AN"` -> all selected CRDs are AN.
If `scope` is `"Upstream"` -> all selected CRDs are Upstream.

---

## Step 2 - Prepare Selected Field Data

> **No new execute call needed.** The data was already extracted in Step 1b.

From `all_extracted`, keep only the CRDs the user selected. Call this `selected_crds`.

Attach the scope to each CRD from `crd_scope_map`. Each record now has a `scope` field: `"Downstream"`, `"AN"`, or `"Upstream"`.

This is the dataset for Steps 3 and 3b.

**CRD sheet layout reference** (for context only - the script handles this automatically):
- CRD number: C8 | Title: C10 | Description: B15 onwards
- Existing Field: C13
- Custom Field Details section: **adaptive** - the script locates the `Custom Field Details` header anywhere in rows 1-25 and reads field attributes (Document, Label, Technical name, Type, Reportable, ERP Importable, ERP Exportable, AN Field) from the adjacent value column. Handles any column offset (I/J, J/K, K/L, etc.). Falls back to K10-K17 if the header is not found.
- Visibility condition: inline in field section (adaptive), then J31 fallback | FMD Name: M10 | Relation Entry: M20
- Each extracted field includes a `_layout_detected` diagnostic field showing which layout was used.

Each element contains:
- `crd_no`, `title`, `description`
- `struck_off`, `cancelled`, `status`
- `existing_field`, `areas_impacted`
- `fields` - array of field definitions
- `fmd`, `relation_entry`
- `scope` - `"Downstream"`, `"AN"`, or `"Upstream"`

---

## Step 3 - Apply Chargeability Rules

**Skip excluded CRDs.** If `struck_off: true` OR `cancelled: true` OR `fields` array is empty, exclude the CRD entirely.

For each remaining CRD in `selected_crds`, route based on `scope`:
- `scope: "Downstream"` -> **Section A** below, using `engagement_type`
- `scope: "AN"` -> **Section B** below
- `scope: "Upstream"` -> **Section C** below

**The chargeable count comes exclusively from these rules.**

---

### Section A - Downstream Rules

#### Section A0 - Guided Buying Check (run BEFORE Organic DD / SAP Store Pack rules)

For each Downstream CRD, perform this check first:

1. **Detect GB requirement:** Check if the CRD `description` explicitly states the customization IS required in Guided Buying. Trigger only on affirmative statements such as:
   - "Customization Required in Guided Buying? Yes"
   - "GB customization required", "required in GB", "must work in GB"

   Do **NOT** trigger on:
   - Warnings like "may not be possible in Guided Buying" or "customizations may not be possible in GB"
   - GB mentioned only in realm names or supplemental references

2. **If GB customization is NOT required:** Skip - proceed to rules below.
3. **If GB customization IS required:** Read `<scratch>/gb-customization-limitations.md`. Match against 14 categories.
4. **If a matching GB limitation is found:** Apply normal rules (count NOT zeroed). Prepend WARNING to Reason.
5. **If no match:** Note in Reason: "GB requirement confirmed - no GB limitation found."

---

#### Organic DD Rules

1. **Normal field** -> **1**
2. **Computed - Display Only** -> **1**
3. **Computed - ERP/AN Export Only** (not visible in UI) -> **0**
4. **Logic field** -> **1**
5. **Logic field - Ease of Access** -> **0**
6. **FMD object** (standalone) -> **1**
7. **FMD - additional attributes on existing FMD** -> **0**
8. **Approval Nodes** -> **0**
9. **OOTB - 2 or more of V/V/E changed** -> **1**
10. **OOTB - only 1 of V/V/E changed** -> **0**
11. **Relation Entry on OOTB field** -> **1**
12. **Relation Entry on custom field** -> **0**
13. **Carry Over** -> **0**
14. **Product Gap** (making OOTB attribute or FMD.ID reportable) -> **0**
15. **Additional Properties** -> **0**
16. **Baseline Code** -> **0**
17. **Same field across multiple document types** -> **1 total** — the field counts as **1 single unit** regardless of how many document types it appears on. Do NOT add Rule 17 on top of another rule (e.g. Rule 1 + Rule 17 ≠ 2). When Rule 17 applies, it IS the total count = 1.

#### SAP Store Pack Rules

1. **Normal field** -> **1**
2. **Computed - Display Only** -> **1**
3. **Computed - ERP/AN Export Only** (existing field) -> **1**; if new field/FMD being created -> **0**
4. **Logic field** -> **1**
5. **Logic field - Ease of Access** -> **0**
6. **FMD object** (standalone) -> **1**
7. **FMD - additional columns on existing FMD** -> **1**
8. **Approval Nodes** -> **1**
9. **OOTB - any attribute update** -> **1 per field**; translations: 5 per field = 1 count
10. **Relation Entry on OOTB field** -> **1**
11. **Relation Entry on existing custom field** -> **1**
12. **Carry Over** -> **0**
13. **Product Gap** -> **1**
14. **Additional Properties** -> **1 per field**; translations: 1 count per 5 per field
15. **Baseline Code** -> **1 per field**
16. **Same field across multiple document types** -> **1 total** — the field counts as **1 single unit** regardless of how many document types it appears on. Do NOT add Rule 16 on top of another rule (e.g. Rule 1 + Rule 16 ≠ 2). When Rule 16 applies, it IS the total count = 1.

If a CRD has ambiguous or incomplete type information, classify conservatively and flag in Reason.

Assign `chargeability_confidence` (0-100) based on how clearly the chargeability rule can be determined from the CRD data. This score reflects **rule certainty only** — not build-readiness.

**Reduces chargeability confidence:**
- Field type missing, `<Set manually>`, or null — rule selection depends on field type
- Document type missing — some rules are per-document-type
- New field vs existing field flag missing or ambiguous
- Multiple competing rules could apply
- Description too vague to infer field type or modification scope

**Does NOT reduce chargeability confidence:**
- Tech lead description blank — standard at chargeability stage; filled in after assessment
- Label missing or not filled
- Technical name incomplete (e.g. `cus_` prefix only)
- FMD name not filled

Scale:
- 90-100%: Field type, doc type, and new/existing flag all clear — single unambiguous rule applies
- 70-89%: Minor ambiguity — one attribute missing but can be reliably inferred from description
- 50-69%: Field type or new/existing flag unclear — multiple rules could apply
- Below 50%: Too many missing data points to determine the applicable rule — flag for manual review

---

#### Section A1 - Realm Multiplier (applies to BOTH Organic DD and SAP Store Pack)

After determining the base chargeable count, apply the realm multiplier.

**If the base chargeable count is 0, skip this step.**

1. **Detect realms** in the CRD `description`:
   - **Test realm** -> count **1** per entry
   - **Supplemental realms** -> count **1** per entry
   - **Production realm** -> count **0** (production is NOT chargeable and does not contribute to the realm multiplier)
2. `realm_count` = Test + Supplemental entries only
3. If `realm_count` >= 1: `final_chargeable = base x realm_count`. If 0: no multiplier.
4. Add realm detail to Reason.

---

### Section B - Ariba Network (AN) Rules

> **Lazy load — only when scope is AN.** `read_file` → `<scratch>/an-chargeability-rules.md`
> Do NOT load this file for Downstream or Upstream runs.

**NOT CUSTOMIZABLE GATE — Run BEFORE any rule scoring. Do not skip.**

Read the full "Not Customizable — AN UI" list in `<scratch>/an-chargeability-rules.md`. Check both field type AND how the customization must behave.

Key categories requiring extra scrutiny:
- **Supplier-only UI customizations**: Cannot implement different UI behaviour per supplier on a single layout. Mark Not Feasible. Add to Reason: "Two-layout + Supplier Group approach may be possible — reassess under Rules 38–40."
- **Font and text formatting**: font-weight, font-size, carriage returns on any text field.
- **Clickable Buttons**: except Add to Header and Line Item Actions on Invoice Input UI.
- **Search Results UI screens**: PO or Invoice search results pages.

If matched: `technically_feasible: "No"`, `chargeable: 0`, `complexity: "N/A"`, both confidence: 100.

For all other AN CRDs, apply:

1. **New Fields - Add Extrinsic** -> **1** per Document Type (not PO UI)
2. **B&I AN Export Fields** -> **0**
3. **Logic Fields (new field)** -> **0**
4. **Logic Fields - Ease of Access** -> **1**
5. **V/V/E on new fields** -> **0**
6. **Customize a Previously Customized Field** -> **1**
7. **Hiding Invoice UI Sections** -> **1** per section
8. **Resequencing Invoice UI Sections** -> **1**
9. **Customize PO Party Info Table** -> **1**
10. **Move/Add/Remove Columns or headings** -> **1**
11. **Warning Messages/Notes (plain text)** -> **1**
12. **Notes displayed as HTML** -> **2**
13. **OOTB Fields - V/V/E** -> **1**
14. **OOTB Fields - Additional Properties** -> **1** (0 if other custs on same field)
15. **OOTB Fields - Choice values (dropdown)** -> **1**
16. **OOTB Fields - Hide field or section** -> **1** (0 if other custs on same field)
17. **Invoice Sub-Type customizations** -> **0**
18. **Customize PO Header Customer Address Info** -> **1**
19. **Hide Purchase Order section on PO UI** -> **1**
20. **Hide PO Header Supplier Address Info** -> **1**
21. **Invoice Add to Header Drop-Down - Hide items** -> **1**
22. **Invoice Line Item Actions Drop-Down - Hide items** -> **2**
23. **Hide Line Item Options pre-fill section** -> **1**
24. **Invoice PDFs - Australia** -> **1**
25. **Invoice PDFs - South Africa** -> **1**
26. **Invoice PDFs - Colombia** -> **1**
27. **Invoice PDFs - all other countries** -> **1**
28. **Invoice ID Length (99.9%)** -> **0**
29. **Invoice ID Validation (.01%)** -> **1**
30. **Order Confirmation ID Length** -> **1**
31. **Legacy PO Verbiage** -> **1**
32. **PO Footer** -> **1**
33. **Tariff Charges - Header (UI Custs)** -> **1**
34. **Tariff Charges - LI (UI Custs)** -> **1**
35. **Tariff Charges - Customer Configuration** -> **0**
36. **Flipping Service PO LI Tax - UI Custs** -> **1**
37. **Flipping Service PO LI Tax - Customer Config** -> **0**
38. **UI Layout - New UI Layout** -> **1** per new layout
39. **UI Layout - Customizations on new layout** -> each chargeable under its own rule; replications count again; note future custs must go to ALL layouts
40. **Supplier Groups - New Supplier Group** -> **1** per group; note conditional feasibility

If a CRD has ambiguous or incomplete type information, classify conservatively and flag in Reason.

Assign `chargeability_confidence` (0-100) using the same scale and criteria as Downstream — rule certainty only, not build-readiness. Tech lead blank, label missing, technical name incomplete do NOT reduce chargeability confidence.

---

### Section C - Upstream Rules (SAP Ariba Sourcing)

> **Lazy load — only when scope is Upstream.** `read_file` → `<scratch>/upstream-chargeability-rules.md`
> Do NOT load this file for Downstream or AN runs.

**Key principle: 1 field = 1 customization** regardless of the number of changes made to that field.

**CRITICAL — V/V/E conditions on NEW fields:**
If a new field is being created and a V/V/E condition is set at the same time, the condition is included in the Create new field count. Total = 1. The "New condition on existing field" rule (+1) applies ONLY to fields that already exist.

#### New Fields

1. **Create new field** -> **1** (includes all settings: label, type, required, reportable, translations, AND any V/V/E conditions set at creation time)
2. **Create new field to replace existing field** -> **2**
3. **Translations for new fields** -> **0**
4. **V/V/E condition set at new field creation time** -> **0**

#### Field Modifications (Custom or OOTB)

4. **Make Field Required** -> **1**
5. **Update Label or Help Tip Text** -> **1**
6. **Label Change** -> **1**
7. **Hide Field** -> **1**
8. **Edits to Advanced Settings** -> **1**
9. **New condition (V/V/E) on EXISTING field** -> **1** (NOT for new fields)
10. **Sync existing field with existing fields in another form** -> **1**
11. **Translations for existing fields** -> **1**

If a CRD has ambiguous or incomplete type information, classify conservatively and flag in Reason.

Assign `chargeability_confidence` (0-100) using the same scale and criteria as Downstream — rule certainty only, not build-readiness. Tech lead blank, label missing, technical name incomplete do NOT reduce chargeability confidence.

**Feasibility note for Upstream:** Flag expressions and conditions as `"Needs Review"` when feasibility cannot be confirmed from the CRD description alone.

For Upstream results: `engagement_type: "Upstream"`, `platform: "SAP Ariba Sourcing"`.

---

## Step 3b - Assess Technical Feasibility

**Technically Feasible:** Yes / No / Needs Review

**Complexity:** Straightforward / Medium / Complex

Assign `feasibility_confidence` (0-100) based on how well the technical feasibility can be assessed from the CRD description and requirement.

**Reduces feasibility confidence:**
- Visibility/validity/editability conditions present with complex or cross-field expressions — expression feasibility is situational
- Condition expression references custom fields that may not yet exist
- Cross-document references (e.g. auto-populate from a different document type)
- Condition expression has typos or syntax errors
- Requirement involves known situationally-feasible scenarios (e.g. Upstream expressions, AN two-layout approaches)

**Does NOT reduce feasibility confidence:**
- Tech lead description blank — standard at chargeability stage; does not reflect on feasibility of the requirement itself
- Simple, well-understood requirements (new field, standard types, no conditions) are feasible regardless of tech lead input

For AN CRDs already marked Not Feasible in the Not Customizable check, skip - values already set.
For Downstream CRDs with a GB warning (Section A0), assess feasibility based on B&I implementation.

---

## Step 4 - Output the Results

> **MANDATORY ON EVERY RUN — no exceptions.** Steps 4a and 4b must always execute after every analysis, whether it is a full multi-CRD run OR a single targeted CRD check. There is no "analysis-only" mode. Output files are always generated.

### 4a - Build the analysis results JSON

Construct the results array from `selected_crds` and write to scratch:

```json
{
  "crd_no": "CRD-01",
  "title": "...",
  "description": "...",
  "scope": "Downstream",
  "field_type": "Normal Custom Field",
  "platform": "OnDemand",
  "engagement_type": "SAP Store Pack",
  "chargeable": 3,
  "technically_feasible": "Yes",
  "complexity": "Straightforward",
  "reason": "SAP Store Pack - Normal field (base: 1). Realm multiplier: 3 (1 Test + 2 Supplemental). Final: 3.",
  "chargeability_confidence": 92,
  "feasibility_confidence": 88
}
```

For AN CRDs: `scope: "AN"`, `engagement_type: "AN"`, `platform: "Ariba Network"`.
For Upstream CRDs: `scope: "Upstream"`, `engagement_type: "Upstream"`, `platform: "SAP Ariba Sourcing"`.

```
write_file -> <scratch>/crd_results.json
```

---

### 4b - Generate file outputs - Excel + HTML in PARALLEL

> **PARALLEL EXECUTION REQUIRED** - emit both calls in a single response turn.

```
execute -> python "<scratch>/export_excel.py"
            "<scratch>/crd_results.json"
            "<path_to_uploaded_file>"
            "<original_name>_Tool_Assessment.xlsx"

execute -> python "<scratch>/generate_html.py"
            "<scratch>/crd_results.json"
            "crd_chargeability_results.html"
```

**What each script produces:**
- `export_excel.py` - Embeds Tool Assessment, Tool Reason, and Scope columns into the CRD Overview sheet.
- `generate_html.py` - Saves a styled `crd_chargeability_results.html` to the working directory.

Wait for both to complete before proceeding.

**After both scripts complete, explicitly state:**
- The exact Excel filename
- The HTML filename: `crd_chargeability_results.html`
- The working directory path

---

### 4c - Render dashboard cards in chat

Using `render_ui`, emit two cards:

1. **KPI card** - five tiles: CRDs Evaluated, Total Chargeable, Avg Confidence, Feasibility, Complexity. If Multiple scope, add 6th tile: Scope Split.

2. **Results table card** - one row per CRD: CRD No., Title, Scope, Field Type, Chargeable, Technically Feasible, Complexity, Reason, Confidence (C: XX% | F: XX%).

---

### 4d - Chat summary

After the cards, output a short **Summary** (4-5 bullets):
- Scope(s) analyzed, engagement type, total chargeable count
- Breakdown by type
- Feasibility highlights: CRDs marked No or Needs Review
- Complexity breakdown
- CRDs with GB warnings or excluded CRDs (if any)

---

## Processing Rules

- Only process sheets named CRD-XX (XX numeric, CRD-01 through CRD-100)
- Skip: CRD Overview, Translations, Interface Changes, Data Source
- **Exclude CRDs entirely when:** `struck_off: true`, `cancelled: true`, or `fields` array is empty
- If the script fails to read a sheet, flag it in the table rather than silently skipping
- Do NOT re-evaluate the existing Chargeable column in the Overview sheet
- **Chargeable count is determined exclusively by the rules in Step 3**
- **Never surface any pre-filled count values from the Excel file in any output**
- **Output generation (Steps 4a + 4b) is mandatory on every analysis run** — single-CRD targeted checks, fast-path runs, and full multi-CRD runs all produce Excel + HTML output files. Never skip output generation.
