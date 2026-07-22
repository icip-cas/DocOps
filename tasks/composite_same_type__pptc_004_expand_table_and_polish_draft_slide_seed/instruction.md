# Task

This task is part of the DocumentBenchmark composite same-type split.

## Inputs

- Source document: `/root/pptc_004_expand_table_and_polish_draft_slide_seed.pptx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This milestone slide needs an extra status field and executive-ready supporting bullets.

## Single-Step Benchmark Instruction

Insert a new `Status` column between `Milestone` and `Due Date` in the table on Slide 1, then rewrite the three draft bullets on that slide into exactly three labeled bullets starting `Vendor status:`, `Install date:`, and `Client update:` in that order.

## Composite Atomic Operations

S4 Table/Sheet ops, C2 Editing

## Composition Pattern

`restructure -> rewrite`

## Atomic Scope

Only the requested table-column insertion and the three targeted bullet rewrites are explicitly requested.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/pptc_004_expand_table_and_polish_draft_slide_seed.pptx`.
- Keep the output format the same as the input document (`.pptx`).
- Focus on the requested composite operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .pptx slide with the expanded table and polished draft bullets.
