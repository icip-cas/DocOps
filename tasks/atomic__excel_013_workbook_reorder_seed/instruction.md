# Task

This task is adapted from the DocumentBenchmark seed `excel_013`.

## Inputs

- Source document: `/root/excel_013_workbook_reorder_seed.xlsx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This workbook contains the right worksheets but they are arranged in an inconvenient order.

## Single-Step Benchmark Instruction

Reorder the worksheets so the tab order becomes: `Instructions` -> `Lookup` -> `Detail_Q1` -> `Detail_Q2` -> `Summary`.

## Atomic Scope

Only worksheet reordering is explicitly requested in this task.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/excel_013_workbook_reorder_seed.xlsx`.
- Keep the output format the same as the input document (`.xlsx`).
- Focus on the requested atomic operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .xlsx workbook with the requested tab order.

## Why This Variant Is Hard

- The current order is plausible enough that a shallow pass may miss the issue
- Several sheet names are similar and easy to misplace
- The task is structural and should not trigger content edits

## Inspired By

- Sorting and rearranging data using Pivot Table (https://www.reddit.com/r/excel/comments/1sn86sl/sorting_and_rearranging_data_using_pivot_table/)
