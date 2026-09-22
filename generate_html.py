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

# Scope detection
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

# Engagement type (from first downstream result)
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

# Scope split KPI tile (only when multiple scopes present)
scope_kpi = ''
if has_multiple:
    parts = []
    if ds_count > 0: parts.append(f'{ds_count} DS')
    if an_count > 0: parts.append(f'{an_count} AN')
    if up_count > 0: parts.append(f'{up_count} UP')
    scope_kpi = f'<div class="kpi"><div class="val">{parts[0]}</div><div class="lbl">Scope Split</div><div class="sub">{" &middot; ".join(parts[1:])}</div></div>'

# Engagement subtitle
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