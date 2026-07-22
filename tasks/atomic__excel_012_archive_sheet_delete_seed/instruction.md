# Task

This task is adapted from the DocumentBenchmark seed `excel_012`.

## Inputs

- Source document: `/root/excel_012_archive_sheet_delete_seed.xlsx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This workbook contains one obsolete sheet that should no longer be present in the deliverable.

## Single-Step Benchmark Instruction

Delete the worksheet named `Archive_2024_DO_NOT_USE` and leave the rest of the workbook unchanged.

## Atomic Scope

Only sheet deletion is explicitly requested in this task.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/excel_012_archive_sheet_delete_seed.xlsx`.
- Keep the output format the same as the input document (`.xlsx`).
- Focus on the requested atomic operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .xlsx workbook with the obsolete sheet removed.

## Why This Variant Is Hard

- The workbook contains multiple tabs and one of them looks plausibly useful unless the name is read carefully
- The instruction is structural, not content-based
- Nothing else in the workbook should be renamed or edited

## Inspired By

- Excel is not saving multiple sheets (https://www.reddit.com/r/excel/comments/1sncbt9/excel_is_not_saving_multiple_sheets/)
