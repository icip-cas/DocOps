# Task

This task is part of a regenerated L3 single-document workflow set.

## Inputs

- Source workbook: `/root/docops_v2_l3_017_xlsx_accreditation_evidence_pack_rebuild_input.xlsx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

Prepare a clean accreditation evidence package workbook from a messy internal review workbook.

## Single-Workbook Workflow Instruction

Revise the source workbook into a public accreditation evidence package. Complete the full workflow in one output `.xlsx`:

- Save the revised workbook to `/root/submission/docops_v2_l3_017_xlsx_accreditation_evidence_pack_rebuild_output.xlsx`.
- Remove all visible `DRAFT`, `PRIVATE`, `INTERNAL`, `scratch`, and confidential review material, including internal-only sheets.
- Rebuild the workbook with sheets in this exact order:
  `Executive Summary`, `Evidence Register`, `Standards Matrix`, `Finding Tracker`, `Owner Workplan`, `Risk Heatmap`, `Citation Cleanup Log`, `Publication Style Guide`, `Appendix Index`.
- Apply the migrated public evidence-pack style:
  header fill `1B4D3E`, white bold header font, Aptos body font, frozen header row, autofilter, public tab colors, date and percent number formats, and a risk heatmap style.
- Apply every correction-log fix:
  `Standard 2.B` to `Standard 2.B.1`, `Acredidation` to `Accreditation`, `Dr. Mina Cho` to `Dr. Mina Zhou`, `Risk score 5` to `Risk score 4`, and publish date `2026-08-01` to `2026-08-15`.
- Rebuild evidence, standards, findings, workplan, risk, citation cleanup, style guide, and appendix tables as public tables.
- Rebuild `Executive Summary`, `Standards Matrix`, and `Owner Workplan` with formulas, not hardcoded totals.
- Exclude board-room dispute, reviewer personal phone, donor sensitivity note, grievance rumor, do-not-publish notes, and internal legal hold text.

## Atomic Scope

Only the requested single-workbook workflow edits are in scope. Preserve public evidence facts, correction facts, formulas, sheet order, tab colors, table style, risk heatmap style, and summary requirements.

## Output Requirements

- Do not modify the source workbook in place.
- Keep the output format the same as the input workbook (`.xlsx`).
- Complete the full workbook workflow, not just local text replacement.

## Expected Output Type

A revised `.xlsx` artifact that completes a high-complexity L3 single-workbook evidence package cleanup, style migration, correction, and formula rebuild workflow.
