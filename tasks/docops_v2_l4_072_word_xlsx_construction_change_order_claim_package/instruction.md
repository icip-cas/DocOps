# Task

Create the final construction change order request and cost/schedule workbook.

## Inputs

- `docops_v2_l4_072_word_xlsx_construction_change_order_claim_package_input.docx`: draft change order request with internal claim strategy
- `contract_change_clause_excerpt.pdf`
- `owner_directive_od014.docx`
- `rfi_221_response.pdf`
- `daily_reports_bundle.pdf`
- `schedule_update.xlsx`
- `weather_log.xlsx`
- `subcontractor_quote_sq17.pdf`
- `cost_ledger.xlsx`
- `internal_claim_strategy_note.pdf`

## Required Outputs

Write both files to `/root/submission`:

- `change_order_request.docx`
- `change_order_cost_schedule_workbook.xlsx`

## Required Work

- Claim only the owner-directed OD-014 / RFI-221 bypass piping and valve relocation change.
- Request exactly `$52,635`: `$47,850` compensable direct cost plus 10% allowable markup.
- Request exactly 3 compensable calendar days, revising substantial completion from `2026-09-15` to `2026-09-18`.
- Exclude weather delivery disruption from the compensable time extension.
- Preserve Excel native tables, formulas, print areas, data validation, conditional formatting, hidden `Internal Strategy`, hidden `Rules`, and defined names.
- The DOCX must include a real TOC, entitlement table, cost/schedule table, highlighted critical path paragraph, evidence boundary, and cross-file reconciliation.
- Do not expose internal claim strategy, 7-day opening position, weather leverage, 18% markup, settlement anchor, walk-away posture, or internal reserve in public outputs.
