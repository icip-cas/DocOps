# Task

This task is part of the DocumentBenchmark composite same-type split.

## Inputs

- Source document: `/root/excelc_009_insert_owner_column_from_lookup_seed.xlsx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This milestone plan is missing owners that can be recovered from a lookup sheet.

## Single-Step Benchmark Instruction

Insert a new `Owner` column between `Milestone` and `Due Date` on `Plan`, then populate it by pulling the matching owners from `Owner Lookup`.

## Composite Atomic Operations

C1 Extraction, S4 Table/Sheet ops

## Composition Pattern

`extract -> restructure`

## Atomic Scope

Only the inserted `Owner` column and the extracted owner values are explicitly requested.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/excelc_009_insert_owner_column_from_lookup_seed.xlsx`.
- Keep the output format the same as the input document (`.xlsx`).
- Focus on the requested composite operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .xlsx workbook with the new `Owner` column populated from lookup data.
