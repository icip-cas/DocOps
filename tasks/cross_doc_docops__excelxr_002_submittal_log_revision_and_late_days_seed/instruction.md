# Task

This task is part of the DocumentBenchmark cross-document document-ops V7 realworld split.

## Goal

Update only the returned submittal rows with the new revision status and resubmit date.

## Primary Input

- `excelxr_002_submittal_log_revision_and_late_days_seed.xlsx`

## Supporting Inputs

- `reviewer_return_extract.pdf`
- `revision_note.docx`
- `owner_map.xlsx`

## Required Edit

Update only the returned submittal rows with the new revision status and resubmit date, insert a `Late Days` formula column immediately after `Due Date`, and highlight rejected rows in the requested color while preserving filters and existing formulas.

## Output

- Write the revised document to `/root/submission/excelxr_002_submittal_log_revision_and_late_days_seed.xlsx`
- Preserve non-target structure and formatting unless the task explicitly asks you to change it.

## Notes

- Keep the template sheets unchanged
- Do not flatten the workbook into plain cells
