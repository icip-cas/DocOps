# Task

This task is part of the DocumentBenchmark composite same-type split.

## Inputs

- Source document: `/root/excelc_007_delete_archive_and_generate_cover_note_seed.xlsx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This workbook still contains an obsolete archive sheet and needs a concise cover note for reviewers.

## Single-Step Benchmark Instruction

Delete the worksheet named `Archive_OLD`, then write exactly two short lines into `Cover!B3`: the first must start with `Supplier:` and the second must start with `Inspection:`.

## Composite Atomic Operations

S1 Insert/Delete, C3 Generation

## Composition Pattern

`delete -> generate`

## Atomic Scope

Only the archive-sheet deletion and the two-sentence cover note are explicitly requested.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/excelc_007_delete_archive_and_generate_cover_note_seed.xlsx`.
- Keep the output format the same as the input document (`.xlsx`).
- Focus on the requested composite operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .xlsx workbook without `Archive_OLD` and with a two-sentence note in `Cover!B3`.
