# SAP Ariba B&I — Technical Feasibility Guide

Source: SAP Ariba Tech Lead Onboarding Guidebook (Internal, 2017) + B&I implementation domain knowledge.

---

## Core Feasibility Principles

### 1. Safe vs Unsafe Fields
- Every field in the Ariba object model is classified as 'safe' or 'unsafe'.
- **Only safe fields can be used for customization.** Unsafe fields are not accessible via Field Configuration.
- If a CRD requires modification of an unsafe field, it is **Not Feasible** as-is. Communicate this to the customer before committing to delivery.

### 2. Customizable vs Non-Customizable Classes
The Ariba class hierarchy has three categories:
- **Fully customizable** — Supports both persisted and derived custom fields. Examples: Requisition, PurchaseOrder, Invoice, ContractRequest, Receipt, InvoiceReconciliation.
- **Derived-only** — Allows derived (computed) fields only. No persistent custom fields can be created on these classes.
- **Non-customizable** — Cannot be modified at all. CRDs targeting these classes are **Not Feasible**.

Common customizable document-level classes:
- `ariba.purchasing.core.Requisition` (Req Header)
- `ariba.purchasing.core.ReqLineItem` (Req Line Item)
- `ariba.purchasing.core.PurchaseOrder` (PO Header)
- `ariba.purchasing.core.POLineItem` (PO Line Item)
- `ariba.invoicing.core.Invoice` (Invoice Header)
- `ariba.invoicing.core.InvoiceLineItem` (Invoice Line Item)
- `ariba.invoicing.core.InvoiceReconciliation` (IR Header)
- `ariba.contract.core.ContractRequest` (Contract Request)
- `ariba.receiving.core.Receipt` (Receipt)

### 3. Field Pool Limits (Persisted Fields)
- Each customizable class has a pre-defined, finite pool of unused fields by data type (String, Boolean, Date, Money, Integer, etc.).
- Example: Requisition Header supports a maximum of 25 String fields.
- If the required field type is exhausted on the target class, the CRD may not be feasible without a workaround (e.g., repurposing a different field type — a Boolean can sometimes be represented as a String Enumeration with fixed values 'Yes'/'No').
- **Derived fields do not consume from this pool** — no limit applies to computed/derived fields.
- Flag as **Needs Review** if there is any risk of hitting field pool limits (especially on frequently customized classes like ReqLineItem or SplitAccounting).

---

## Complexity Indicators by Field Type

| Field Type | Typical Complexity | Notes |
|---|---|---|
| OOTB – Labels/Translations/Tooltips only | Straightforward | No AML required; handled entirely via Field Configuration UI |
| OOTB – Single V/V/E change | Straightforward | Simple condition in Field Configuration |
| OOTB – 2+ V/V/E changes | Medium | Multiple conditions; requires careful condition logic and testing |
| Normal (cus_ field, standard document class) | Straightforward | Standard add-field workflow via Field Configuration |
| Computed / Derived field (simple expression) | Straightforward | AML expression is straightforward; no storage required |
| Computed with cross-document or conditional logic | Medium | AML expressions with conditionals or cross-object field references |
| FMD object (new or existing) | Medium | Requires FMD setup + optional AML lookup expression |
| FMD within FMD | Complex | Nested FMD lookup; non-trivial AML pattern with careful null handling |
| Relation Entry (OOTB field) | Medium | Chargeable; requires Relation Entry configuration |
| Relation Entry (custom field) | Medium | Not chargeable but requires AML-based validation logic |
| Logic – simple single-condition approval | Medium | Single approval rule condition |
| Logic – multi-condition approval or invoice exception | Complex | Nested and/or conditions; accounting-level access; careful null/boolean handling |
| Accounting field validation (SplitAccounting, SAP realm) | Complex | SAP realm-specific rules; multi-ERP considerations apply |
| AN extrinsic (new extrinsic field mapping) | Medium | Requires AN-side extrinsic definition + B&I export mapping |
| AN logic (value pass-through, e.g., PO → SES) | Complex | Cross-document value pull; requires AML + AN-side understanding |
| AN OOTB V/V/E (any one property) | Medium | Different rule from OnDemand — any single V/V/E = chargeable on AN |
| Multi-ERP or multi-child-site customization | Complex | Per-site counting and configuration overhead; realm-specific behaviour |
| Customization on derived-only class | Complex | Workaround required; no persistent field possible |
| Carry-over of existing field to another document | Straightforward | Same requirement, slight adaptation; no new design needed |

---

## Common Not-Feasible Scenarios

- Requirement targets a non-customizable class
- Requirement explicitly involves an 'unsafe' field
- Requirement asks for real-time external callout during document approval without a supported integration mechanism
- Customization on OrderConfirmation or ASN for AN (not in Downstream Deployment scope)
- Requirement expects platform behaviour that does not exist in the current Ariba B&I release

---

## Common Needs-Review Scenarios

- CRD description is vague about which Ariba document or class the field belongs to
- Multiple document types mentioned without a clear carry-over indication
- Requirement may require a field type (e.g., Boolean, Date) that could be near the pool limit on a commonly-used class
- Requirement involves AN but lacks clarity on whether a B&I export mapping already exists
- Logic field requirement where it is unclear whether this qualifies as business logic (chargeable) or ease-of-access (not chargeable)
- Technical name not provided — cannot confirm if field is custom (`cus_` prefix) or OOTB
- Requirement references a feature introduced in a specific quarterly release — compatibility with the customer's version needs checking

---

## Best Practice Notes (from Ariba Tech Lead Onboarding Guide)

- Always declare new fields on the **common super-class** wherever possible to avoid duplicate code across sub-classes.
- **Internationalization**: All custom field labels must have translations defined. Missing translations are a common post-go-live issue.
- **Relation Entry Validation**: Relation entries on custom fields should include AML-based null-safe validation to avoid runtime errors.
- **FMD Lookup**: FMD lookups in AML must handle the case where the FMD object does not exist — always include null guards.
- **Accounting Validations (SAP realm)**: Fields on SplitAccounting behave differently in multi-ERP setups. Test against all configured ERP systems.
- **Null and Boolean handling**: Ariba AML has non-obvious null coercion behaviour. Boolean fields stored as String ('Yes'/'No') require explicit null checks before comparisons.
- **Power of Inheritance**: A field declared on a parent class is available across all child classes — use this to avoid redundant customization across similar document types.
