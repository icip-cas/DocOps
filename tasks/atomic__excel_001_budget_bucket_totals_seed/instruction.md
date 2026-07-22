# Task

This task is adapted from the DocumentBenchmark seed `excel_001`.

## Inputs

- Source document: `/root/excel_001_budget_bucket_totals_seed.xlsx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This workbook tracks costs with a code in one column and needs totals split into Maintenance and Operations budgets.

## Single-Step Benchmark Instruction

Using formulas, calculate the correct Maintenance total and Operations total in the summary area based on the cost-code mapping in the `Code Map` sheet.

## Atomic Scope

Only the summary computation is explicitly requested in this task.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/excel_001_budget_bucket_totals_seed.xlsx`.
- Keep the output format the same as the input document (`.xlsx`).
- Focus on the requested atomic operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A repaired .xlsx workbook with correct formula-driven totals.

## Latent Issues Intentionally Left In The Seed

- The Budget Bucket column in the main table is blank
- The summary panel formatting is uneven
- The workbook layout is cluttered and could be improved, but that is not the explicit request

## Inspired By

- Directing costs to multiple different totals based on category (https://www.reddit.com/r/excel/comments/1snym9r/directing_costs_to_multiple_different_totals/)
