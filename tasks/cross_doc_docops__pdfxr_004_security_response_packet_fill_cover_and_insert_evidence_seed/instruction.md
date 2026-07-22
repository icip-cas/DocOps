# Task

This task is part of the DocumentBenchmark cross-document document-ops V7 realworld split.

## Goal

Fill the cover placeholders from the answered questionnaire.

## Primary Input

- `pdfxr_004_security_response_packet_fill_cover_and_insert_evidence_seed.pdf`

## Supporting Inputs

- `cover_values.xlsx`
- `response_rules.docx`
- `customer_summary.pdf`
- `bookmark_plan.docx`

## Required Edit

Fill the cover placeholders from the answered questionnaire, insert the required evidence pages in the specified section order, and keep the legal disclaimer and existing approved appendix pages unchanged.

## Output

- Write the revised document to `/root/submission/pdfxr_004_security_response_packet_fill_cover_and_insert_evidence_seed.pdf`
- Preserve non-target structure and formatting unless the task explicitly asks you to change it.

## Notes

- Preserve the image-bearing evidence page
- Do not flatten the packet into regenerated plain pages
