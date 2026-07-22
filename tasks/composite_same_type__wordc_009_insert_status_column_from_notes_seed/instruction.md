# Task

This task is part of the DocumentBenchmark composite same-type split.

## Inputs

- Source document: `/root/wordc_009_insert_status_column_from_notes_seed.docx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This action table is missing status values that are stated in the notes below it.

## Single-Step Benchmark Instruction

Insert a new `Status` column between `Owner` and `Due Date`, then populate each row using the statuses stated in the note bullets under the table.

## Composite Atomic Operations

C1 Extraction, S4 Table/Sheet ops

## Composition Pattern

`extract -> restructure`

## Atomic Scope

Only the inserted `Status` column and extracted status values are explicitly requested.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/wordc_009_insert_status_column_from_notes_seed.docx`.
- Keep the output format the same as the input document (`.docx`).
- Focus on the requested composite operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .docx document with a populated `Status` column added to the table.
