# Task

This task is adapted from the DocumentBenchmark seed `excel_009`.

## Inputs

- Source document: `/root/excel_009_body_style_consistency_seed.xlsx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This workbook contains a candidate table where the body rows use inconsistent fonts, emphasis, and fills.

## Single-Step Benchmark Instruction

On the `Hiring Tracker` sheet, make rows `4:10` use one consistent body style that matches row `4`, while leaving the title row, header row, and note in row `13` unchanged.

## Atomic Scope

Only formatting consistency is explicitly requested in this task.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/excel_009_body_style_consistency_seed.xlsx`.
- Keep the output format the same as the input document (`.xlsx`).
- Focus on the requested atomic operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .xlsx workbook with normalized body-row styling.

## Why This Variant Is Hard

- Several rows look intentionally emphasized even though they are inconsistent noise
- The instruction references a style exemplar rather than listing every formatting attribute
- Non-target rows should remain untouched

## Inspired By

- Cells in table won't update format (https://www.reddit.com/r/excel/comments/1snfjl6/cells_in_table_wont_update_format/)
