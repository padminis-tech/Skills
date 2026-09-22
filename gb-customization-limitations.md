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
