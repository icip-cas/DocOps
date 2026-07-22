# Task

This task is part of the DocumentBenchmark composite same-type split.

## Inputs

- Source document: `/root/pdfc_003_fill_summary_box_and_apply_theme_seed.pdf`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This report has an empty summary box and inconsistent accent colors.

## Single-Step Benchmark Instruction

Replace the placeholder text in the `Management Summary` box on page 1 with exactly three short lines starting `Supplier:`, `Inspection:`, and `Staffing:` in that order, then retheme the title bar and callout boxes to a cohesive blue-gray palette.

## Composite Atomic Operations

C3 Generation, F4 Theme transfer

## Composition Pattern

`generate -> retheme`

## Atomic Scope

Only the three constrained summary bullets and the requested blue-gray theme update are explicitly requested.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/pdfc_003_fill_summary_box_and_apply_theme_seed.pdf`.
- Keep the output format the same as the input document (`.pdf`).
- Focus on the requested composite operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .pdf document with a filled summary box and a cohesive blue-gray theme.
