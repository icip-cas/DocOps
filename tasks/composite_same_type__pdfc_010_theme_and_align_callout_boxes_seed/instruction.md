# Task

This task is part of the DocumentBenchmark composite same-type split.

## Inputs

- Source document: `/root/pdfc_010_theme_and_align_callout_boxes_seed.pdf`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This one-page PDF has mismatched accent colors and uneven callout alignment.

## Single-Step Benchmark Instruction

Apply a cohesive blue-gray theme to the callout boxes and align the three callout boxes into one evenly spaced horizontal row beneath the title.

## Composite Atomic Operations

F4 Theme transfer, F3 Layout control

## Composition Pattern

`retheme -> realign`

## Atomic Scope

Only the callout theming and alignment are explicitly requested.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/pdfc_010_theme_and_align_callout_boxes_seed.pdf`.
- Keep the output format the same as the input document (`.pdf`).
- Focus on the requested composite operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .pdf page with aligned callout boxes and a cohesive blue-gray theme.
