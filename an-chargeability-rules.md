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
| Adding Tariff Charges to Invoice | Add Invoice Header Level Tariff Charges (UI Custs) | **1** | Only UI customization needed is to remove "Allowance" from the "Add to Header" drop-down. Customer must use OOTB Allowances & Charges feature — cannot add "Tariff Charge" directly to drop-down menus. B&I supports Allowances & Charges OOTB and will only require minor customizations. Tariff Charges must be reflected in Invoice Totals, which is why the Allowances & Charges feature is required (a plain Extrinsic would not be included in totals). |
| Adding Tariff Charges to Invoice | Add Invoice LI Tariff Charges (UI Custs) | **1** | Only UI customization needed is to remove "Allowance" from the "Line Item Actions" drop-down. If other customizations already prevent suppliers from adding Header level charges, this is the only customization needed. |
| Adding Tariff Charges to Invoice | Add Invoice LI Tariff Charges (Customer) | **0** | Customer task only — not a UI customization: (1) Enable "Allow suppliers to add allowances and charges to invoice" transaction rule; (2) Configure a custom Charge for Tariffs in Invoice Transaction rules. Step 2 is dependent on Step 1 being completed first. |
| Flipping Service PO LI Tax to Invoice | UI Customizations | **1** | Default PO-Flip OOTB functionality does not flip LI tax from Composite Service POs (POs with Parent and Child lines) to the Invoice — LI Tax on Parent lines is not flipped. Special Invoice UI customization required to flip LI tax from Parent lines into Invoice child lines. |
| Flipping Service PO LI Tax to Invoice | Customer Configurations/Tasks | **0** | Customer task only — not a UI customization: (1) Enable "Copy tax from purchase order to standard invoice" transaction rule; (2) Open support ticket to enable Feature Toggle CSC-16389 on both Test and Production ANIDs (typically 2 separate tickets). |
| UI Layouts | New UI Layout | **1** | One count per new UI Layout created. Required when different sets of suppliers need different UI behaviour (e.g. different dropdown choices per supplier type). Each new layout is a standalone build. |
| UI Layouts | Customizations deployed on a new layout | **(per rule)** | Each customization on a new layout is independently chargeable under its own applicable rule — including customizations replicated from an existing layout. Replications count again because they are built fresh on the new layout. Important: any future Invoice UI customization must be applied to ALL layouts in use, permanently multiplying maintenance effort — flag this in the Reason. |
| Supplier Groups | New Supplier Group | **1** | Creating a new supplier group to associate a UI Layout with a specific subset of suppliers = 1. Suppliers not assigned to any group automatically receive the default UI Layout. Feasibility is conditional: if any supplier needs both UI behaviours (e.g. both Sales Tax and VAT), the supplier-group approach is not feasible — a supplier can only belong to one group at a time. |

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
- **Clickable Buttons** — except Add to Header and Line Item Actions on the Invoice Input UI (includes Save, Exit, Update, Next, Previous, Submit, Add, Delete, Create Invoice, and all other buttons)
- **Standard PDF direct customizations** — PDFs are indirectly customized when the related UI screen is customized
- **Auto-generated email notifications** — emails are not CommAuto document UI screens
- **SBN Reports** — all reports have pre-designed formats and cannot be customized
- **Custom Mappings in Managed Gateway for Spend and Network** — PO and OC/ASN/SES/INV Custom Mappings are not included in UI Customizations
- **Multi-columnar table format for multiple sets of Extrinsic values** — not supported
- **Supplier-only UI customizations** — AN UI customizations are static and apply identically to all suppliers. If a requirement needs different UI behaviour for different suppliers (e.g. different dropdown choices, different field visibility depending on which supplier is viewing the document), it cannot be implemented on a single UI Layout. Mark as Not Feasible for the original scope. Note: a two-layout + supplier group approach may be an alternative, but this transforms the requirement into a different, more expensive scope — reassess chargeability using the UI Layout and Supplier Group rules above. Conditional feasibility applies: the two-layout approach is not viable if any supplier needs both UI behaviours simultaneously, since a supplier can only belong to one group.
- **OC Line Item Non-Extrinsic field customizations** — only Extrinsic customizations supported on OC Line Item UI
- **ASN UI sections with known limitations** — certain sections do not support customizations

---

## Multi-Site Notes

AN chargeability is assessed per CRD. Multi-site rules follow the same principles as Downstream — apply site multipliers manually after getting per-CRD counts.