# Task

This task is adapted from the DocumentBenchmark seed `excel_003`.

## Inputs

- Source document: `/root/excel_003_risk_highlighting_seed.xlsx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This tracker sheet has risk scores, but urgent items do not stand out visually.

## Single-Step Benchmark Instruction

Visually highlight all rows where the `Risk Score` is greater than or equal to 85.

## Atomic Scope

Only highlighting/emphasis is explicitly requested in this task.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/excel_003_risk_highlighting_seed.xlsx`.
- Keep the output format the same as the input document (`.xlsx`).
- Focus on the requested atomic operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .xlsx workbook where high-risk rows are clearly highlighted.

## Latent Issues Intentionally Left In The Seed

- Column widths and layout are rough
- Status logic could be improved
- Dates are not visually prioritized, but that is not the explicit request

## Inspired By

- Conditional formatting based on a different cell's value (https://www.reddit.com/r/excel/comments/1sndxcd/conditional_formatting_based_on_a_different_cells/)
