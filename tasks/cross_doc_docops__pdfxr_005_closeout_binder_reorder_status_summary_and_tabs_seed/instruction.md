# Task

This task is part of the DocumentBenchmark cross-document document-ops V7 realworld split.

## Goal

Reorder the binder into the requested closeout sequence.

## Primary Input

- `pdfxr_005_closeout_binder_reorder_status_summary_and_tabs_seed.pdf`

## Supporting Inputs

- `closeout_index.xlsx`
- `warranty_note.docx`
- `bookmark_plan.docx`

## Required Edit

Reorder the binder into the requested closeout sequence, refresh the cover status summary with exact counts, and rebuild bookmarks for O&M, Warranties, Punch List, and Training while preserving non-target manual pages.

## Output

- Write the revised document to `/root/submission/pdfxr_005_closeout_binder_reorder_status_summary_and_tabs_seed.pdf`
- Preserve non-target structure and formatting unless the task explicitly asks you to change it.

## Notes

- Keep the image-bearing certificate and O&M pages
- Do not regenerate the binder as plain text pages
