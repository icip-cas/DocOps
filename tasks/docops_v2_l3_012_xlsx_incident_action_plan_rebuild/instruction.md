# Task

This task is part of a regenerated L3 single-document workflow set.

## Inputs

- Source workbook: `/root/docops_v2_l3_012_xlsx_incident_action_plan_rebuild_input.xlsx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

Prepare a clean public Incident Action Plan workbook for a citywide heat emergency from a messy internal planning workbook.

## Single-Workbook Workflow Instruction

Revise the source workbook into a public IAP workbook. Complete the full workflow in one output `.xlsx`:

- Save the revised workbook to `/root/submission/docops_v2_l3_012_xlsx_incident_action_plan_rebuild_output.xlsx`.
- Remove all visible `DRAFT`, `PRIVATE`, `INTERNAL`, `scratch`, and confidential planning material, including internal-only sheets.
- Rebuild the workbook with sheets in this exact order:
  `Summary`, `Incident Objectives`, `Organization Assignments`, `Branch Assignments`, `Communications Plan`, `Medical Plan`, `Correction Log`, `Style Guide`.
- Apply the migrated public IAP style from the style reference:
  header fill `7030A0`, white bold header font, Calibri body font, frozen header row, and autofilter on tabular sheets.
- Rebuild `Incident Objectives`, `Organization Assignments`, `Branch Assignments`, `Communications Plan`, and `Medical Plan` as clean public tables.
- Apply every correction-log fix:
  `North Cooling Cneter` to `North Cooling Center`, `Marcos Silva` to `Marisol Silva`, `Medcial` to `Medical`, `TAC-6` to `TAC-7`, and ambulance staging `1800` to `1900`.
- Rebuild `Branch Assignments` with formulas for `Total Staff`, not hardcoded totals.
- Rebuild `Summary` with formulas linked to `Branch Assignments`, `Medical Plan`, and `Communications Plan`.
- Keep the public style guide sheet, but remove all private implementation notes.
- Exclude vendor liability, mayor backchannel, personal phone, volunteer medical rumor, shelter shortage rumor, do-not-release notes, and staffing grievance text.

## Atomic Scope

Only the requested single-workbook workflow edits are in scope. Preserve public incident facts, correction facts, formulas, sheet order, and style migration requirements.

## Output Requirements

- Do not modify the source workbook in place.
- Keep the output format the same as the input workbook (`.xlsx`).
- Complete the full workbook workflow, not just one local sheet edit.

## Expected Output Type

A revised `.xlsx` artifact that completes a complex L3 single-workbook IAP cleanup, correction, formula rebuild, and style migration workflow.
