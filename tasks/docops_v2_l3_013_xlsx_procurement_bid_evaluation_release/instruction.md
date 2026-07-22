# Task

This task is part of a regenerated L3 single-document workflow set.

## Inputs

- Source workbook: `/root/docops_v2_l3_013_xlsx_procurement_bid_evaluation_release_input.xlsx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

Prepare a clean public procurement bid evaluation workbook from a messy internal evaluation workbook.

## Required Sequential Workflow

This task must be completed in order:

1. Remove internal-only sheets and confidential evaluator notes.
2. Apply correction-log fixes before rebuilding public scoring tables.
3. Rebuild the vendor register first, because later data validation must reference it.
4. Create the `VendorIDs` named range from the corrected vendor IDs.
5. Rebuild scoring, price normalization, and award recommendation formulas.
6. Add data validation dropdowns that use the `VendorIDs` named range.
7. Add public debrief hyperlinks.
8. Apply the required public workbook style.

## Single-Workbook Workflow Instruction

Revise the source workbook into a public bid evaluation release workbook. Complete the full workflow in one output `.xlsx`:

- Save the revised workbook to `/root/submission/docops_v2_l3_013_xlsx_procurement_bid_evaluation_release_output.xlsx`.
- Remove all visible `DRAFT`, `PRIVATE`, `INTERNAL`, `scratch`, and confidential evaluator material, including internal-only sheets.
- Rebuild the workbook with sheets in this exact order:
  `Evaluation Summary`, `Vendor Register`, `Requirements Scoring`, `Price Normalization`, `Conflict-of-Interest Log`, `Award Recommendation`, `Debrief Register`, `Data Validation Guide`, `Publication Style Guide`, `Appendix Index`.
- Apply all correction-log fixes:
  `VND-03A` to `VND-003`, `Northstar Analytcs` to `Northstar Analytics`, evaluator `Leah Moor` to `Leah Moore`, price `$1,240,000` to `$1,204,000`, and status `Responsiv` to `Responsive`.

## Required Non-Text Features

- Create named range `VendorIDs` pointing to `'Vendor Register'!$A$2:$A$6`.
- Add data validation list `=VendorIDs` on `Requirements Scoring!B2:B16`.
- Add data validation list `=VendorIDs` on `Conflict-of-Interest Log!B2:B6`.
- Add `mailto:` hyperlinks in `Debrief Register!D2:D6`.

## Required Style Migration

- Use sheet tab colors exactly as defined in the public style guide.
- Use header fill `4F2D7F` on every public table.
- Use white bold header font on every public table.
- Use Aptos as the body font.
- Freeze the header row on all public tabular sheets.
- Enable autofilter on all public tabular sheets.
- Use currency format `$#,##0` on all public price columns.
- Use percent format `0%` on all score-weight columns.

## Content Requirements

- Rebuild vendor register, requirements scoring, price normalization, conflict log, award recommendation, debrief register, validation guide, style guide, and appendix index as public tables.
- Exclude evaluator disagreement, personal phone, vendor protest strategy, legal hold, do-not-release notes, scoring rumor, and private negotiation text.

## Output Requirements

- Do not modify the source workbook in place.
- Keep the output format the same as the input workbook (`.xlsx`).
- Complete the full ordered workflow, not just local text replacement.

## Expected Output Type

A revised `.xlsx` artifact that completes a strict-order L3 single-workbook procurement evaluation workflow with named ranges, data validation, hyperlinks, formulas, cleanup, and explicit style migration.
