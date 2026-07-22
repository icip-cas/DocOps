# Task

This task is part of the DocumentBenchmark composite same-type split.

## Inputs

- Source document: `/root/pptc_006_rewrite_bullets_and_standardize_text_style_seed.pptx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This status deck has rough draft bullets and inconsistent text styling.

## Single-Step Benchmark Instruction

Rewrite the three draft bullets on Slide 2 so they sound executive-ready, then standardize the bullet text style on that slide so all three bullets match.

## Composite Atomic Operations

C2 Editing, F1 Style consistency

## Composition Pattern

`rewrite -> restyle`

## Atomic Scope

Only the three bullet rewrites and their style normalization on Slide 2 are explicitly requested.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/pptc_006_rewrite_bullets_and_standardize_text_style_seed.pptx`.
- Keep the output format the same as the input document (`.pptx`).
- Focus on the requested composite operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .pptx deck with polished bullets and consistent bullet styling on Slide 2.
