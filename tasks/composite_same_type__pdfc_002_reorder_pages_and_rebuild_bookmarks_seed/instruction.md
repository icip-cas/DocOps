# Task

This task is part of the DocumentBenchmark composite same-type split.

## Inputs

- Source document: `/root/pdfc_002_reorder_pages_and_rebuild_bookmarks_seed.pdf`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This report is out of order and the bookmark tree should match the corrected sequence.

## Single-Step Benchmark Instruction

Reorder the pages so the report reads `Scope -> Findings -> Recommendations -> Appendix`, then rebuild the bookmarks so `Site A` and `Site B` are nested under `Findings`.

## Composite Atomic Operations

S2 Reorder, S3 Hierarchy editing

## Composition Pattern

`reorder -> rebuild-structure`

## Atomic Scope

Only the page reorder and bookmark hierarchy rebuild are explicitly requested.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/pdfc_002_reorder_pages_and_rebuild_bookmarks_seed.pdf`.
- Keep the output format the same as the input document (`.pdf`).
- Focus on the requested composite operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .pdf document with corrected page order and bookmark hierarchy.
