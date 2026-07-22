# Task

This task is part of the DocumentBenchmark composite same-type split.

## Inputs

- Source document: `/root/excelc_010_compute_risk_flag_and_highlight_rows_seed.xlsx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This tracker needs a computed risk flag and visible emphasis on the flagged items.

## Single-Step Benchmark Instruction

Create a `Risk Flag` column on `Tracker` using formulas that mark rows as `Critical` when score is at least `85`, then highlight the full flagged rows.

## Composite Atomic Operations

C4 Computation, F2 Highlighting

## Composition Pattern

`compute -> highlight`

## Atomic Scope

Only the formula-based flag column and highlighting of the flagged rows are explicitly requested.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/excelc_010_compute_risk_flag_and_highlight_rows_seed.xlsx`.
- Keep the output format the same as the input document (`.xlsx`).
- Focus on the requested composite operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .xlsx workbook with a computed `Risk Flag` column and highlighted critical rows.
