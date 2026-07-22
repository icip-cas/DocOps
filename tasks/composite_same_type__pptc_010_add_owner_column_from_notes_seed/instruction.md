# Task

This task is part of the DocumentBenchmark composite same-type split.

## Inputs

- Source document: `/root/pptc_010_add_owner_column_from_notes_seed.pptx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This milestone table is missing owner values that can be recovered from the notes panel.

## Single-Step Benchmark Instruction

Insert a new `Owner` column into the milestone table on Slide 1, then populate it using the owner names stated in the notes panel on the same slide.

## Composite Atomic Operations

C1 Extraction, S4 Table/Sheet ops

## Composition Pattern

`extract -> restructure`

## Atomic Scope

Only the inserted `Owner` column and extracted owner values are explicitly requested.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/pptc_010_add_owner_column_from_notes_seed.pptx`.
- Keep the output format the same as the input document (`.pptx`).
- Focus on the requested composite operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .pptx slide with a populated `Owner` column added to the milestone table.
