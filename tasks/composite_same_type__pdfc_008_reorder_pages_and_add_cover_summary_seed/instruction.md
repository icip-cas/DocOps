# Task

This task is part of the DocumentBenchmark composite same-type split.

## Inputs

- Source document: `/root/pdfc_008_reorder_pages_and_add_cover_summary_seed.pdf`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This report should be reordered and open with a concise management summary box.

## Single-Step Benchmark Instruction

Reorder the pages to `Scope -> Findings -> Recommendations -> Appendix`, then fill the cover summary box with exactly three short lines starting `Permit:`, `Supplier:`, and `Staffing:` in that order.

## Composite Atomic Operations

S2 Reorder, C3 Generation

## Composition Pattern

`reorder -> generate`

## Atomic Scope

Only the page reorder and the three-line cover summary are explicitly requested.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/pdfc_008_reorder_pages_and_add_cover_summary_seed.pdf`.
- Keep the output format the same as the input document (`.pdf`).
- Focus on the requested composite operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .pdf report with corrected page order and a populated cover summary box.
