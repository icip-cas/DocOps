# Task

This task is part of a regenerated L3 single-document workflow set.

## Inputs

- Source presentation: `/root/docops_v2_l3_022_pptx_port_cyber_ttx_public_deck_input.pptx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

Prepare a clean public after-action briefing deck from a messy internal port cyber tabletop exercise presentation.

## Single-Presentation Workflow Instruction

Revise the source presentation into a polished public after-action deck. Complete the full workflow in one output `.pptx`:

- Save the revised presentation to `/root/submission/docops_v2_l3_022_pptx_port_cyber_ttx_public_deck_output.pptx`.
- Remove all visible and speaker-note `DRAFT`, `PRIVATE`, `INTERNAL`, `scratch`, and confidential cyber-response material.
- Remove internal-only slides and rebuild the deck with exactly these eight public slides in this order:
  1. `Port Azure Cyber Tabletop Exercise: Public After-Action Briefing`
  2. `Exercise Scope`
  3. `Incident Timeline`
  4. `Decision Log`
  5. `Operational Impact Matrix`
  6. `Recovery Priorities`
  7. `Public Communications`
  8. `After-Action Commitments`
- Migrate the approved public deck style from the style reference slide:
  title slide background `0B1F33`, content slide background `F5F7FA`, accent color `00A896`, Trebuchet MS bold titles, and public footer text.
- Apply all correction-log fixes:
  `Terminal 3` to `Terminal 2`, `Crane PLC outage 74 minutes` to `47 minutes`, `Public update 11:20` to `11:05`, `Recovery target 36 hours` to `24 hours`, and `CISO Dana Rhee` to `CISO Dana Reed`.
- Rebuild the incident timeline, decision log, operational impact matrix, recovery priorities, public communications, and after-action commitments as clean public slides.
- Keep concise public speaker notes on every slide, but remove confidential notes and personal contact details.
- Exclude ransomware attribution, insurer call, unpaid invoice, board blame, personal phone, do-not-disclose notes, and law-enforcement-sensitive details.

## Atomic Scope

Only the requested single-presentation workflow edits are in scope. Preserve public tabletop facts, correction facts, slide order, speaker-note cleanup, and style migration requirements.

## Output Requirements

- Do not modify the source presentation in place.
- Keep the output format the same as the input presentation (`.pptx`).
- Complete the full deck workflow, not just one local slide edit.

## Expected Output Type

A revised `.pptx` artifact that completes a complex L3 single-presentation public after-action briefing workflow.
