# Task

Create the vendor B-notice action memo and backup withholding tracker workbook.

## Inputs

- `docops_v2_l4_076_word_xlsx_vendor_bnotice_backup_withholding_packet_input.docx`
- `irs_cp2100a_notice.pdf`
- `irs_backup_withholding_rules.pdf`
- `corrected_w9_packets.docx`
- `vendor_master_and_payment_run.xlsx`
- `internal_ap_risk_note.docx`

## Required Outputs

Write both files to `/root/submission`:

- `vendor_bnotice_action_memo.docx`
- `backup_withholding_tracker.xlsx`

## Required Work

- Process the CP2100A first B-notice workflow.
- Identify corrected W-9 vendors and vendors still missing valid TIN documentation.
- Apply the 24% backup withholding rule only to reportable payments that require it.
- Update the vendor master correction plan and 1099 evidence trail.
- Build a workbook with native Excel tables, formulas, data validation, print areas, hidden raw/rules sheets, and defined names.
- Keep internal AP risk notes and vendor-risk commentary out of public-facing outputs.
