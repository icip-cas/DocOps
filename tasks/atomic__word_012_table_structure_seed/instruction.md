# Task

This task is adapted from the DocumentBenchmark seed `word_012`.

## Inputs

- Source document: `/root/word_012_table_structure_seed.docx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This inspection table is missing a Status column.

## Single-Step Benchmark Instruction

Insert a new Status column between Owner and Due Date while preserving the existing rows.

## Atomic Scope

Only the table structure change is explicitly requested in this task.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/word_012_table_structure_seed.docx`.
- Keep the output format the same as the input document (`.docx`).
- Focus on the requested atomic operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .docx document with the requested table modification.

## Why This Variant Is Hard

- The change is structural rather than textual
- Existing data must stay aligned with the right rows
- The insertion point matters, so appending at the end is not sufficient
