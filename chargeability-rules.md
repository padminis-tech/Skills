# SAP Ariba Chargeability Rules — Buying and Invoicing Custom Fields

Source: Procurement - Chargeable Custom Fields Delivery Guidelines - Buying and Invoicing, Catalogs, Invoice Mgmt (2021)

**Scope: SAP Ariba APPS (Downstream / OnDemand Solution) only.**

---

## OnDemand Solution (B&I) — Organic DD vs SAP Store Pack

| Field Type | Field Characteristic | Organic DD | Organic DD Notes | SAP Store Pack | SAP Store Pack Notes |
|---|---|:---:|---|:---:|---|
| Normal Fields | Any | **1** | Fields on a document/approvable or master data object. Count does not vary with effort. | **1** | Same as Organic DD. |
| Computed Fields | Display Only (visible in UI) | **1** | Computed fields added to UI view | **1** | Same as Organic DD. |
| Computed Fields | ERP/AN Exports Only (not UI visible) | **0** | Integration fields not visible in UI | **1** | Counted as 1 if it's an existing field. If a new field or FMD is being created, do not double count — the new field/FMD count suffices (count 0 here). |
| Logic Fields | Business logic (approvals, invoice exceptions) | **1** | Fields for approvals, custom invoice exceptions | **1** | Same as Organic DD. |
| Logic Fields | Ease of Access | **0** | Purely technical fields at accounting/line level to compute or access header expressions. No direct customer requirement; internal debugging. | **0** | Same as Organic DD. These fields do not deliver the CRD requested. |
| FMD | Object (standalone or as right field for filtering/validation) | **1** | The FMD object itself. Normal Field of type FMD = 1 for the normal field; the FMD object itself does not add another count in that case. | **1** | Same as Organic DD. |
| FMD | Additional attributes/custom columns on an existing FMD | **0** | Extra columns on an existing FMD object. Example: FMD with 20 custom fields = 1 total (not 21). | **1** | Adding columns to an existing FMD = 1 (regardless of how many columns). If those newly added columns are used individually on other documents/transactions, count each separately per the applicable field type rule. |
| Approval Nodes | Any | **0** | Approval nodes for customer requirements | **1** | Even if it is an update only. |
| OOTB Fields | 2 or more of Visibility / Validity / Editability changed | **1** | Customization touches 2+ of V/V/E | **1** | Any attribute update = 1 per field. |
| OOTB Fields | Only 1 of Visibility / Validity / Editability changed | **0** | Only one property modified | **1** | Any attribute update including a single V/V/E change, template update, or translations = 1 per field. 5 translations per field = 1 count. |
| Relation Entries | For OOTB fields | **1** | Relation entries on standard OOTB fields | **1** | Same as Organic DD. |
| Relation Entries | For custom fields | **0** | The custom field itself accounts for chargeability; relation entry on it adds nothing. | **1** | Relation entries for already existing custom fields = 1. |
| Carry Over | Same field carried across approvables/documents with similar functionality | **0** | Avoids double counting for the same requirement. | **0** | Same as Organic DD. |
| Product Gap | Making OOTB attribute or FMD.ID reportable (not customer-specific) | **0** | Not customer-specific. Examples: Supplier.Name, FMD.ID | **1** | Making any existing field (custom or OOTB) reportable = 1. |
| Additional Properties | Labels, translations, tooltips | **0** | Field characteristics only | **1** | For existing fields = 1 per field. Translations: 1 count per 5 translations per field. |
| Baseline Code | Uncomment/comment template configuration | **0** | — | **1** | Uncommenting/commenting baseline code = 1 per field. |
| Same Field – Multiple Document Types | Same custom field required on multiple document types | **1** | Counted as a single chargeable unit regardless of how many document types it spans. | **1** | Each unique custom field = 1, regardless of how many document types it is deployed on. |

---

## Multi-Solution / Multi-Child Site Rules

Chargeability is counted **per solution** and **per child site**:
- 5 free customizations per solution per child site (embedded services)
- Customer on B&I + Inv-Mgmt with 2 child sites = 5 x 2 (solutions) x 2 (child sites) = 20 free customizations
- If a custom field is deployed to 2 child sites = counted as 2

Note: The skill does NOT calculate multi-site totals — apply these rules manually after getting per-CRD counts.

---

## Scope Note

These guidelines cover SAP Ariba APPS (Downstream / OnDemand B&I) only.
- Customizations on OrderConfirmation and ASN do NOT fall under Downstream Deployment (DD) scope.
- AN (Ariba Network) chargeability guidelines are covered separately in `references/an-chargeability-rules.md`.
