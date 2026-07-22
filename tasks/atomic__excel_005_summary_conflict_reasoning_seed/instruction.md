# Task

This task is adapted from the DocumentBenchmark seed `excel_005`.

## Inputs

- Source document: `/root/excel_005_summary_conflict_reasoning_seed.xlsx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This workbook contains a short summary sheet that may not match the actual regional performance data.

## Single-Step Benchmark Instruction

Identify which note on the `Ops Summary` sheet is inconsistent or incomplete when compared with the `Monthly Data` sheet.

Reply using exactly this 4-line template:

`Note: 1`
`North: actual 93 vs target 95`
`East: actual 83 vs target 88`
`Why: North is not the only region below target because East is also below target.`

## Atomic Scope

Only reasoning over consistency between sheets is explicitly requested in this task.

## Output Requirements

- Do not modify the source document in place.
- Write your final answer to `/root/submission/final_answer.txt`.
- Keep the answer concise, directly responsive, and grounded in the provided document.

## Expected Output Type

A concise, template-constrained explanation of the inconsistent note.

## Why This Variant Is Hard

- Several summary notes are correct, so the task is not simple error spotting by suspicion alone
- The incorrect statement is partially true but still misleading
- The workbook requires cross-sheet comparison, not just local reading

## Inspired By

- Tips on my Dashboard Summary and Sales Tracker (https://www.reddit.com/r/excel/comments/1sle1k0/tips_on_my_dashboard_summary_and_sales_tracker/)
