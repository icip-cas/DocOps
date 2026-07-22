# Task

This task is part of a regenerated L3 single-document workflow set.

## Inputs

- Source workbook: `/root/docops_v2_l3_015_xlsx_reimbursement_workbook_formula_repair_input.xlsx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

Prepare a clean public reimbursement repair workbook from a messy quarterly grant workbook with data-entry, category, formula, and audit-trail errors.

## Required Sequential Workflow

This task must be completed in order:

1. Remove draft, private, and obsolete scratch sheets.
2. Apply the correction log before rebuilding formulas or summaries.
3. Repair transaction dates, categories, eligibility flags, reimbursement formulas, and audit notes.
4. Rebuild the category rollup and dashboard from the corrected transaction register.
5. Rebuild the exception log from the corrected transactions.
6. Add required data validation, conditional formatting, audit comments, sheet protection, and workbook structure protection.
7. Save the clean workbook as a new `.xlsx` without modifying the source workbook.

## Single-Workbook Workflow Instruction

Revise the source workbook into one clean public `.xlsx` workbook:

- Save the revised workbook to `/root/submission/docops_v2_l3_015_xlsx_reimbursement_workbook_formula_repair_output.xlsx`.
- Remove visible `DRAFT`, `PRIVATE`, scratch, obsolete, external-link, and uncertainty material.
- Rebuild the workbook with exactly these five sheets in this order:
  1. `Repair Dashboard`
  2. `Transaction Register`
  3. `Category Rollup`
  4. `Exception Log`
  5. `Control Lists`

## Required Corrections

- `TX-004` category must be `Training`, not `Travel`.
- `TX-006` category must be `Lobbying` and eligibility must be `No`.
- `TX-008` service date must be `2026-06-30`, not `2026-07-01`.
- `TX-007` must remain a separate valid transaction after duplicate review.
- Reimbursement formulas must return the amount only when `Eligible` is `Yes`; otherwise `0`.
- Quarter formulas must derive the quarter from `Service Date`.
- Audit status formulas must mark ineligible rows as `Exception`, corrected rows as `Corrected`, and untouched eligible rows as `Ready`.

## Required Workbook Features

- Include formulas, not hardcoded results, for quarter, reimbursement amount, audit status, dashboard metrics, and category rollup.
- Add audit comments to the corrected exception note cells.
- Add list validation for corrected category and eligibility columns.
- Add conditional formatting for exception status and reimbursable amounts.
- Add a dashboard bar chart from `Category Rollup`.
- Protect `Repair Dashboard` and `Control Lists`, and protect workbook structure.

## Required Style

- Use header fill `305496` with white bold header text.
- Use body font `Aptos` size `11`.
- Use currency number format `$#,##0;($#,##0);-`.
- Freeze panes as specified in the metadata.

## Output Requirements

- Do not modify the source workbook in place.
- Keep the output format the same as the input workbook (`.xlsx`).
- Complete the full ordered workflow, not just local cell replacement.

## Expected Output Type

A revised `.xlsx` artifact that completes a strict-order L3 single-workbook repair workflow with corrected data, formula repair, audit comments, exception logging, dashboard rollup, data validation, conditional formatting, protection, and style migration.
