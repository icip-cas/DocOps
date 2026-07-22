# Task

This task is part of the DocumentBenchmark composite same-type split.

## Inputs

- Source document: `/root/pptc_008_fix_outline_and_align_cards_seed.pptx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This workstream slide has both incorrect bullet nesting and uneven card alignment.

## Single-Step Benchmark Instruction

Fix the bullet hierarchy so task-level lines are nested under the correct workstreams, then align the three workstream cards into one even row.

## Composite Atomic Operations

S3 Hierarchy editing, F3 Layout control

## Composition Pattern

`repair-hierarchy -> align`

## Atomic Scope

Only the bullet hierarchy repair and card alignment on the target slide are explicitly requested.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/pptc_008_fix_outline_and_align_cards_seed.pptx`.
- Keep the output format the same as the input document (`.pptx`).
- Focus on the requested composite operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .pptx slide with corrected bullet nesting and aligned workstream cards.
