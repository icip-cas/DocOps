# Task

This task is part of the DocumentBenchmark L4 manual expansion set.

## Goal

Prepare a final owner closeout binder and a matching warranty follow-up tracker.

## Inputs

- `docops_v2_l4_052_pdf_xlsx_closeout_binder_and_warranty_tracker_input.pdf`: draft closeout binder PDF
- `punchlist_log.xlsx`: punchlist status and owners
- `warranty_register.xlsx`: warranty status and owners
- `photo_appendix.pdf`: closeout photo evidence
- `superintendent_closeout_note.docx`: authoritative release notes and exclusions

## Required Outputs

Write both files to `/root/submission`:

- `docops_v2_l4_052_pdf_xlsx_closeout_binder_and_warranty_tracker_input.pdf`
- `warranty_followup_tracker.xlsx`

## Required Edit

Remove internal-only pages and draft-only language. Carry only open punchlist and unresolved warranty items:

- Include PL-04 and PL-07.
- Exclude closed PL-11.
- Include W-18 and W-29.
- Exclude received W-22.

The PDF cover must show 2 open punchlist items, 2 open warranty items, and critical follow-up `PL-04 and W-18`.

## Preservation Requirements

- Rebuild PDF bookmarks for the final binder.
- Create a native Excel table named `closeout_followup_tracker`.
- Preserve formulas in `Follow-up Tracker!I2:I5` and `Release Summary!B3:B5`.
- Preserve data validation on `Follow-up Tracker!G2:G5`.
- Preserve hidden `Rules` sheet and `status_choices` defined name.
- PDF binder rows and XLSX tracker rows must agree.
