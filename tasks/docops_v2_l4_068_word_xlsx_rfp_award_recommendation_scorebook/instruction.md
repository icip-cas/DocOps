# Task

Create the final RFP award recommendation memo and evaluation scorebook.

## Inputs

- `docops_v2_l4_068_word_xlsx_rfp_award_recommendation_scorebook_input.docx`: draft award memo with incorrect lowest-price recommendation and internal notes
- `raw_scoring_export.xlsx`: raw vendor scores and prices
- `rfp_evaluation_rules.pdf`: weights, mandatory gate, price normalization, and public-boundary rules
- `vendor_proposal_excerpts.pdf`: proposal excerpts and mandatory-gate evidence
- `evaluator_conflict_note.docx`: conflict note requiring Taylor's CivicStack score exclusion

## Required Outputs

Write both files to `/root/submission`:

- `award_recommendation_memo.docx`
- `evaluation_scorebook.xlsx`

## Required Work

- Recommend CivicStack, not Northstar Data.
- Treat Northstar Data as ineligible because it failed the mandatory data-residency gate.
- Build a native Excel scorebook with Eligibility, Score Matrix, Price Analysis, Public Debrief, hidden Weights, hidden Conflict Log, and hidden Rules sheets.
- Preserve formulas, native tables, print areas, data validation, conditional formatting, and defined names.
- The award memo must include a real TOC, ranked vendor table, highlighted conflict-control note, and best-value tradeoff rationale.
- Do not disclose evaluator names outside the conflict-control note, proprietary OCR routing, discount floors, or raw evaluator comments in the public memo.
