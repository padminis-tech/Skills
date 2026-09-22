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
| Field Modifications | New condition (Visibility, Editability, Validity) on EXISTING field | **1** | Adding a new V/V/E condition to an already-existing field = 1 customization. This rule applies ONLY to existing fields — NOT to new fields where the condition is set at creation time (see New Fields table above). |
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