# Task

This task is part of the DocumentBenchmark composite same-type split.

## Inputs

- Source document: `/root/wordc_004_insert_status_column_and_formalize_notes_seed.docx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This action tracker section needs a richer table and more professional notes.

## Single-Step Benchmark Instruction

Insert a new `Status` column between `Owner` and `Due Date` in the action table, then rewrite only the two draft note paragraphs into a labeled pair where the first starts `Schedule risk:` and the second starts `Client update:`.

## Composite Atomic Operations

S4 Table/Sheet ops, C2 Editing

## Composition Pattern

`restructure -> rewrite`

## Atomic Scope

Only the requested table-column insertion and the two targeted note rewrites are explicitly requested.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/wordc_004_insert_status_column_and_formalize_notes_seed.docx`.
- Keep the output format the same as the input document (`.docx`).
- Focus on the requested composite operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .docx document with the new table column and polished draft notes.
