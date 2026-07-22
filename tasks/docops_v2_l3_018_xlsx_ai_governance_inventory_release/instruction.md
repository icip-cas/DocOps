# Task

This task is part of a regenerated L3 single-document workflow set.

## Inputs

- Source workbook: `/root/docops_v2_l3_018_xlsx_ai_governance_inventory_release_input.xlsx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

Prepare a clean public AI governance inventory workbook from a messy internal AI risk review workbook.

## Single-Workbook Workflow Instruction

Revise the source workbook into a public AI governance inventory and risk workbook. Complete the full workflow in one output `.xlsx`:

- Save the revised workbook to `/root/submission/docops_v2_l3_018_xlsx_ai_governance_inventory_release_output.xlsx`.
- Remove all visible `DRAFT`, `PRIVATE`, `INTERNAL`, `scratch`, and confidential model-review material, including internal-only sheets.
- Rebuild the workbook with sheets in this exact order:
  `Governance Summary`, `AI System Inventory`, `RMF Function Matrix`, `Risk Assessment Register`, `Evaluation Evidence`, `Incident Log`, `Human Oversight Plan`, `Public Transparency Register`, `Publication Style Guide`, `Appendix Index`.
- Apply all correction-log fixes:
  `Goven` to `Govern`, `Mappping` to `Mapping`, `Chatbot Triage v0.8` to `Chatbot Triage v1.0`, bias evaluation date `2026-05-01` to `2026-06-10`, and owner `Ravi Shah` to `Ravi Sharma`.

## Required Style Migration

The output must apply these explicit workbook style requirements:

- Use sheet tab colors exactly as defined in the public style guide.
- Use header fill `243B53` on every public table.
- Use white bold header font on every public table.
- Use Aptos as the body font.
- Freeze the header row on all public tabular sheets.
- Enable autofilter on all public tabular sheets.
- Use date format `yyyy-mm-dd` on all public date columns.
- Use percent format `0%` on all public percent columns.
- Use the risk register heatmap colors: `High = C00000`, `Medium = F4B183`, `Low = 70AD47`.
- Use formulas, not hardcoded values, for summary totals, average evaluation coverage, and risk counts.

## Content Requirements

- Rebuild AI system inventory, RMF function matrix, risk assessment register, evaluation evidence, incident log, human oversight plan, transparency register, style guide, and appendix index as public tables.
- Exclude prompt injection transcript, vendor security exception, employee performance note, personal phone, legal hold, do-not-publish note, model failure rumor, and private benchmark score.

## Atomic Scope

Only the requested single-workbook workflow edits are in scope. Preserve public AI governance facts, correction facts, formulas, sheet order, tab colors, number formats, heatmap style, and table style requirements.

## Output Requirements

- Do not modify the source workbook in place.
- Keep the output format the same as the input workbook (`.xlsx`).
- Complete the full workbook workflow, not just local text replacement.

## Expected Output Type

A revised `.xlsx` artifact that completes a high-complexity L3 single-workbook AI governance inventory cleanup, correction, formula, and explicit style migration workflow.
