# Task

This task is part of the DocumentBenchmark cross-document document-ops V7 realworld split.

## Goal

Convert the sectioned request list into the folder-import workbook.

## Primary Input

- `excelxr_005_due_diligence_folder_import_numbering_and_owners_seed.xlsx`

## Supporting Inputs

- `diligence_request_extract.pdf`
- `folder_note.docx`
- `team_map.xlsx`

## Required Edit

Fill `Folder Import` with the exact hierarchy marker / folder title / responsible team rows, set every imported row to `Missing`, and fill the launch metadata cells on `Cover`.

## Output

- Write the revised document to `/root/submission/excelxr_005_due_diligence_folder_import_numbering_and_owners_seed.xlsx`
- Preserve non-target structure and formatting unless the task explicitly asks you to change it.

## Notes

- Keep `Folder Template` unchanged
- Do not replace summary formulas with hardcoded values
