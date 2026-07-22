# Task

Create the payroll garnishment setup memo and garnishment calculation tracker workbook.

## Inputs

- `docops_v2_l4_048_pdf_docx_xlsx_payroll_garnishment_priority_packet_input.docx`
- `employee_profile.docx`
- `child_support_iwo.pdf`
- `consumer_garnishment_order.pdf`
- `payroll_register.xlsx`
- `ccpa_rules_excerpt.pdf`
- `state_priority_rules.pdf`
- `internal_hr_risk_note.docx`

## Required Outputs

Write both files to `/root/submission`:

- `payroll_garnishment_setup_memo.docx`
- `garnishment_calculation_tracker.xlsx`

## Required Work

- Calculate disposable earnings using legally required deductions only.
- Do not subtract voluntary 401(k), medical premium, or gym deductions from disposable earnings.
- Apply the child-support withholding order before the ordinary consumer garnishment.
- Determine the correct child-support cap and ordinary consumer garnishment cap.
- Decide the exact biweekly withholding amounts for `CS-22-1189` and `CG-7741`.
- Build a tracker with native Excel tables, formulas, data validation, print areas, hidden raw/rules sheets, and defined names.
- Prepare court, employee, and payroll communication controls.
- Keep internal HR/legal commentary, manager debt commentary, and improper shortcuts out of public-facing outputs.
