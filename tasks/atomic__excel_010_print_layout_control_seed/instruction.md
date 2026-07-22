# Task

This task is adapted from the DocumentBenchmark seed `excel_010`.

## Inputs

- Source document: `/root/excel_010_print_layout_control_seed.xlsx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This workbook contains a weekly planning sheet that is readable on screen but poorly configured for print layout.

## Single-Step Benchmark Instruction

Adjust the `Weekly Plan` sheet so the table prints cleanly on one landscape page with the header row repeated and no obviously clipped wrapped text, without changing any cell values.

## Atomic Scope

Only page/layout control is explicitly requested in this task.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/excel_010_print_layout_control_seed.xlsx`.
- Keep the output format the same as the input document (`.xlsx`).
- Focus on the requested atomic operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .xlsx workbook with improved print layout settings.

## Why This Variant Is Hard

- The problem is caused by several interacting layout settings rather than one broken cell
- Wrapped narrative text makes row height and column width tradeoffs nontrivial
- The request is print-oriented, not merely cosmetic on-screen formatting

## Inspired By

- Making a fillable form with dropdowns and sections that can be added and moved easily (https://www.reddit.com/r/excel/comments/1sn0f6w/making_a_fillable_form_with_dropdowns_and/)
