# Task

Create the final retail lease abstract memo and obligation calendar.

## Inputs

- `docops_v2_l4_070_word_xlsx_retail_lease_abstract_obligation_calendar_input.docx`: draft lease abstract with superseded renewal date and internal strategy
- `executed_lease_excerpt.pdf`: base lease clauses for term, rent, CAM, insurance, and reporting
- `amendment_2_renewal_cam.docx`: controlling renewal option and CAM objection amendment
- `landlord_cam_statement.xlsx`: landlord 2025 CAM statement
- `internal_renewal_strategy_note.pdf`: internal-only renewal strategy

## Required Outputs

Write both files to `/root/submission`:

- `lease_abstract_memo.docx`
- `lease_obligation_calendar.xlsx`

## Required Work

- Correct the renewal notice window to 2028-07-01 through 2028-09-30.
- Calendar exactly five public obligations: two renewal dates, CAM objection deadline, insurance certificate, and gross sales report.
- Calculate CAM cap review and flag the 2025 controllable CAM over-cap amount.
- Preserve Excel native tables, formulas, print areas, data validation, conditional formatting, hidden `Internal Strategy`, hidden `Rules`, and defined names.
- The DOCX memo must include a real TOC, critical-date table, highlighted CAM objection paragraph, evidence boundary, and cross-file reconciliation.
- Do not expose INT-900, the renewal floor, walk-away strategy, option leverage, or landlord pressure plan in public sheets or the memo.
