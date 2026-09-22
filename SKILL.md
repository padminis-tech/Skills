---
name: ariba-crd-chargeability
description: >-
  Analyzes an uploaded Ariba CRD (Customization Requirement Description) Excel tracker to assess the chargeability of each custom field requirement. Supports three scopes: SAP Ariba APPS (Downstream), Ariba Network (AN), and Upstream (SAP Ariba Sourcing). For Downstream: supports Organic DD and SAP Store Pack engagement types. For AN and Upstream: single rule set each. Outputs dashboard cards in chat, a downloadable Excel file, and a styled HTML dashboard. Use this skill whenever the user uploads a CRD Excel file and asks to check chargeability, count chargeable fields, assess feasibility, or review CRD field types. Trigger phrases: "check CRD chargeability", "analyze CRD", "count chargeable fields", "CRD feasibility check", "check my CRD tracker", "Organic DD chargeability", "SAP Store Pack chargeability", "AN chargeability", "Ariba Network CRD", "check AN customizations", "upstream chargeability", "sourcing CRD", "upstream CRD".
allowed-tools: execute write_file read_file render_ui
metadata:
  author: Ariba Technology Consultant
  version: 4.8.3
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
execute -> python "<skill_dir>/scripts/extract_crd.py" "<path_to_uploaded_file>" "all"
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
3. **If GB customization IS required:** Read `references/gb-customization-limitations.md`. Match against 14 categories.
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

> Full rule reference: `references/an-chargeability-rules.md`

**NOT CUSTOMIZABLE GATE — Run BEFORE any rule scoring. Do not skip.**

Read the full "Not Customizable — AN UI" list in `references/an-chargeability-rules.md`. Check both field type AND how the customization must behave.

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

> Full rule reference: `references/upstream-chargeability-rules.md`

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
execute -> python "<skill_dir>/scripts/export_excel.py"
            "<scratch>/crd_results.json"
            "<path_to_uploaded_file>"
            "<original_name>_Tool_Assessment.xlsx"

execute -> python "<skill_dir>/scripts/generate_html.py"
            "<scratch>/crd_results.json"
            "crd_chargeability_results.html"
```

**What each script produces:**
- `export_excel.py` - Embeds Tool Assessment, Tool Reason, and Scope columns into the CRD Overview sheet.
- `generate_html.py` - Saves a styled `crd_chargeability_results.html` to the working directory.

Both scripts are bundled in the skill's `scripts/` folder - **do NOT rewrite them**.

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