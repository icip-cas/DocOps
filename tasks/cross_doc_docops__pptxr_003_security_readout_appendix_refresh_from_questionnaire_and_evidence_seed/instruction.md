# Task

This task is part of the DocumentBenchmark cross-document document-ops V7 realworld split.

## Goal

Update the customer answer grid.

## Primary Input

- `pptxr_003_security_readout_appendix_refresh_from_questionnaire_and_evidence_seed.pptx`

## Supporting Inputs

- `summary_note.docx`
- `control_status.xlsx`
- `quote_note.pdf`

## Required Edit

Update the customer answer grid, duplicate the appendix reference slide to create `Evidence Highlights`, fill three evidence bullets and IDs, and preserve the theme and reference-slide geometry.

The final live slide order must be: `Policy Readout`, `Summary`, `Risks`, `Evidence Divider`, `Evidence Highlights`, `Customer Answer Grid`, `Closing`, followed by the two unchanged locked reference slides.

## Output

- Write the revised document to `/root/submission/pptxr_003_security_readout_appendix_refresh_from_questionnaire_and_evidence_seed.pptx`
- Preserve non-target structure and formatting unless the task explicitly asks you to change it.

## Notes

- Keep both reference slides unchanged
- Do not rebuild the inserted slides with different layouts
