# Task

This task is adapted from the DocumentBenchmark seed `excel_002`.

## Inputs

- Source document: `/root/excel_002_last_nonblank_lookup_seed.xlsx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This workbook stores monthly values by year and product, with many blanks. The user wants one result cell that returns the latest available value for the selected year and selected product.

## Single-Step Benchmark Instruction

Fill the result cell on the `Selector` sheet with a formula that returns the last non-blank monthly value for the chosen product and chosen year.

## Atomic Scope

Only the selector formula is explicitly requested in this task.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/excel_002_last_nonblank_lookup_seed.xlsx`.
- Keep the output format the same as the input document (`.xlsx`).
- Focus on the requested atomic operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A repaired .xlsx workbook where the selector cell returns the correct last non-blank value.

## Latent Issues Intentionally Left In The Seed

- The data layout is awkward and repeated horizontally
- Styling is inconsistent across sheets
- The workbook could be normalized into a cleaner table, but that is not the explicit request

## Inspired By

- Two-way XLOOKUP to return the last non-blank value at a row-column intersection (https://www.reddit.com/r/excel/comments/1snc945/twoway_xlookup_to_return_the_last_nonblank_value/)
