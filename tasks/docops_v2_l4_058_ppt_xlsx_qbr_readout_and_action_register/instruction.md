# Task

This task is part of the DocumentBenchmark L4 manual expansion set.

## Goal

Refresh the QBR renewal readout and create a matching action register workbook.

## Inputs

- `docops_v2_l4_058_ppt_xlsx_qbr_readout_and_action_register_input.pptx`: draft deck
- `renewal_risk_source.xlsx`: authoritative risk source
- `account_owner_notes.docx`: canonical owner names and required actions
- `finance_exception_notice.pdf`: authoritative conflict overrides

## Required Outputs

Write both files to `/root/submission`:

- `docops_v2_l4_058_ppt_xlsx_qbr_readout_and_action_register_input.pptx`
- `renewal_action_register.xlsx`

## Required Edit

Use the finance exception notice as authoritative over the draft deck:

- Northstar Bank must be `Watch`, not `Hold`.
- Zenith Foods ARR must be `$1.40M`, not `$1.20M`.

Remove the scratch slide. Reorder the live deck to:

1. QBR Renewal Readout
2. Executive Summary
3. Renewal Risk Table
4. Action Register
5. Risk Callout
6. Reference - Risk Callout
7. Reference - Action Table

The PPTX risk table, PPTX action table, and XLSX action register must agree exactly on account, owner, action, and status.

## Preservation Requirements

- Preserve both reference slides and their text.
- The XLSX action register must use a native table named `renewal_action_register`.
- Preserve the hidden `Rules` sheet.
- Preserve the `risk_status_choices` defined name.
- Preserve data validation on `Action Register!E2:E4`.
- Keep live formulas in `Action Register!G2:G4`; do not hardcode `OK`.
