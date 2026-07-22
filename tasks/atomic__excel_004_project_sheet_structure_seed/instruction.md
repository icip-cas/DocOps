# Task

This task is adapted from the DocumentBenchmark seed `excel_004`.

## Inputs

- Source document: `/root/excel_004_project_sheet_structure_seed.xlsx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This workbook has multiple project sheets with the same task-table structure, but each one is missing a `Priority` column.

## Single-Step Benchmark Instruction

Insert a new `Priority` column between `Task` and `Owner` on both `Project_A` and `Project_B` while preserving the existing row data.

## Atomic Scope

Only the table/sheet structure change is explicitly requested in this task.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/excel_004_project_sheet_structure_seed.xlsx`.
- Keep the output format the same as the input document (`.xlsx`).
- Focus on the requested atomic operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .xlsx workbook with the requested structural change on both project sheets.

## Latent Issues Intentionally Left In The Seed

- The workbook has no consistent table styling
- There is no master roll-up logic yet
- Dates and statuses are not validated, but that is not the explicit request

## Inspired By

- Pulling data out of multiple tables across sheets to a master table? (https://www.reddit.com/r/excel/comments/1smraf4/pulling_data_out_of_multiple_tables_across_sheets/)
- Inserting new, whilst copying adjacent protected formulas (https://www.reddit.com/r/excel/comments/1smbh65/inserting_new_whilst_copying_adjacent_protected/)
