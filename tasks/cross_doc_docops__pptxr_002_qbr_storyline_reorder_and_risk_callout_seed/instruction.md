# Task

This task is part of the DocumentBenchmark cross-document document-ops V7 realworld split.

## Goal

Reorder the storyline slides into the requested sequence.

## Primary Input

- `pptxr_002_qbr_storyline_reorder_and_risk_callout_seed.pptx`

## Supporting Inputs

- `risk_note.docx`
- `issue_log.xlsx`
- `storyline_order.docx`

## Required Edit

Reorder the live slides into the requested sequence, replace the `Open Risks` table from the issue log, insert a bounded `Risk Callout` slide by duplicating the locked reference layout, and keep the reference slide unchanged.

## Output

- Write the revised document to `/root/submission/pptxr_002_qbr_storyline_reorder_and_risk_callout_seed.pptx`
- Preserve non-target structure and formatting unless the task explicitly asks you to change it.

## Notes

- Keep the reference slide unchanged
- Do not rebuild the inserted slide with a different geometry
