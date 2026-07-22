# Task

Create the utility rebate M&V summary deck and dynamic measure calculation tracker.

## Inputs

- `docops_v2_l4_061_pptx_xlsx_utility_rebate_mv_packet_input.docx`
- `energy_audit_report.docx`
- `brightgrid_rebate_guide.pdf`
- `vendor_quote_and_cutsheets.pdf`
- `baseline_utility_usage.xlsx`
- `internal_pricing_strategy.docx`

## Required Outputs

Write both files to `/root/submission`:

- `utility_rebate_mv_summary.pptx`
- `rebate_measure_calculation_tracker.xlsx`

## Required Work

- Identify eligible measures and exclude ineligible measures.
- Calculate annual kWh savings, installed cost, gross incentive, 50% cost cap, approved incentive, and net cost.
- Apply the lighting `$0.08/kWh` rule, VFD `$120/hp` rule, and 50% installed-cost cap.
- Exclude garage stairwell controls-only sensors and east wing code-required BAS scheduling.
- Build a workbook with native Excel tables, formulas, data validation, print areas, hidden raw/rules sheets, and defined names.
- Build a six-slide PPTX summary for the owner/utility reviewer.
- Do not expose private pricing strategy, invoice-splitting, inflated hours, or margin notes in public outputs.
