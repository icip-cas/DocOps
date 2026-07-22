# Task

This task is part of a regenerated L3 single-document workflow set.

## Inputs

- Source workbook: `/root/docops_v2_l3_014_xlsx_disaster_housing_grant_public_release_input.xlsx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

Prepare a clean public disaster housing repair grant release workbook from a messy internal intake workbook.

## Required Sequential Workflow

This task must be completed in order:

1. Remove private intake, duplicate-scratch, and legal-hold material before creating the public release workbook.
2. Deduplicate applications by `Application ID`, keeping the latest `Inspection Date`.
3. Apply the correction bulletin before calculating eligibility:
   - `H-104` county must be `North Fork`.
   - `H-109` damage category must be `Major`.
   - `H-112` income band must be `Low`.
   - Assistance caps are `Minor = 12000`, `Moderate = 18000`, `Major = 28000`, and `Destroyed = 42000`.
4. Rebuild the public eligibility register with formulas for eligible amount, priority score, and review status.
5. Build county rollup formulas and a dashboard chart from the corrected eligibility register.
6. Build an appeal queue from applications with appeal flag `Yes`.
7. Apply the required public release styles, freeze panes, conditional formatting, sheet protection, and workbook structure protection.

## Single-Workbook Workflow Instruction

Revise the source workbook into one clean public `.xlsx` workbook:

- Save the revised workbook to `/root/submission/docops_v2_l3_014_xlsx_disaster_housing_grant_public_release_output.xlsx`.
- Remove all visible `DRAFT`, `PRIVATE`, `legal hold`, PII, raw-intake, duplicate-scratch, and unverified hardship material.
- Remove applicant names, personal phone numbers, private notes, and private-only tabs.
- Rebuild the workbook with exactly these five sheets in this order:
  1. `Funding Dashboard`
  2. `Eligibility Register`
  3. `County Rollup`
  4. `Appeal Queue`
  5. `Publication Controls`

## Required Calculation Rules

- `Eligible Amount` must use an Excel formula equivalent to:
  `MAX(0, MIN(Repair Estimate - Insurance Proceeds, assistance cap for corrected Damage Category))`.
- `Priority Score` must use an Excel formula equivalent to:
  vulnerable household `Yes` adds `20`; damage adds `40` for `Destroyed`, `30` for `Major`, `20` for `Moderate`, and `10` for `Minor`; income band adds `20` for `Low`, `10` for `Moderate`, and `0` for `High`.
- `Review Status` must use an Excel formula equivalent to:
  `Senior Review` when eligible amount is at least `20000`; otherwise `Appeal Review` when appeal flag is `Yes`; otherwise `Ready`.
- `County Rollup` and `Funding Dashboard` must use Excel formulas that reference the corrected public register.

## Required Dashboard and Chart Feature

- `Funding Dashboard` must include formulas for total applications, total eligible amount, senior review count, appeal review count, and average priority score.
- `Funding Dashboard` must include a bar chart using county names from `County Rollup!A2:A5` and eligible amount values from `County Rollup!C2:C5`.

## Required Style and Workbook Controls

- Use header fill color `1F4E79` with white bold header text on all public tables.
- Use body font `Aptos` size `11`.
- Use currency number format `$#,##0;($#,##0);-` for public currency fields.
- Use `A2` freeze panes on all public table sheets and `A4` freeze panes on `Funding Dashboard`.
- Apply conditional formatting to highlight `Eligibility Register!K2:K13` when priority score is at least `70`.
- Apply conditional formatting to highlight `Eligibility Register!J2:J13` when eligible amount is greater than `0`.
- Protect `Funding Dashboard` and `Publication Controls`.
- Protect workbook structure.
- Use the required tab colors from the metadata.

## Output Requirements

- Do not modify the source workbook in place.
- Keep the output format the same as the input workbook (`.xlsx`).
- Complete the full ordered workflow, not just local text replacement.

## Expected Output Type

A revised `.xlsx` artifact that completes a strict-order L3 single-workbook public grant release workflow with deduplication, correction-log application, formulas, charting, conditional formatting, privacy cleanup, and workbook protection.
