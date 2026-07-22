# Vendor Compliance Exception Register Repair

You are given a single Excel workbook:

`docops_v2_l3_019_xlsx_vendor_compliance_exception_register_repair_input.xlsx`

The workbook is a messy internal vendor compliance dump. It contains draft-only sheets, personal contact details, named individual owners, inconsistent statuses, stale renewal dates, and unstructured exception notes. Repair it into a public-safe compliance control workbook.

## Required Workflow

Complete the repair in this order:

1. Remove draft-only sheets and all personal contact details.
2. Normalize vendor records into a clean intake table using the public contact mailbox `compliance-review@city.example`.
3. Convert unstructured exceptions into an owner-group based exception register. Do not use named individual owners.
4. Build a renewal calendar with formula-driven days-until-expiry and bucket fields.
5. Build a dashboard with formula-driven KPI cells and a bar chart for open exceptions by risk.
6. Add an audit log that describes repair categories without repeating personal names, personal phone numbers, or draft-only source phrases.
7. Apply workbook controls:
   - Exact sheet order: `Dashboard`, `Clean Intake`, `Exception Register`, `Renewal Calendar`, `Audit Log`.
   - Excel table objects on all data sheets.
   - Data validation for risk and status fields.
   - Conditional formatting for high risk, open status, and expiring renewals.
   - Comments explaining key formula/control assumptions.
   - Protected sheets and locked workbook structure.
   - Header fill `1F2937`, header font `FFFFFF`, body font `Arial` size `10`.
8. Remove all restricted or incorrect source remnants:
   - `DRAFT`
   - `PRIVATE contact`
   - `Ava Stone`
   - `555-0142`
   - `do not share`
   - `personal email`
   - `staff-only`
   - `unreviewed vendor`
   - `compliance-old@city.example`
   - `Named individual owner`
   - `TBD`

Save the completed workbook as:

`/root/submission/docops_v2_l3_019_xlsx_vendor_compliance_exception_register_repair_output.xlsx`
