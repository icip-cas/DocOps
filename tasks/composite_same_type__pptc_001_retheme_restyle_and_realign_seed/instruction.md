# Task

This task is part of the DocumentBenchmark composite same-type split.

## Inputs

- Source document: `/root/pptc_001_retheme_restyle_and_realign_seed.pptx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This deck needs a coherent theme, consistent typography, and aligned workstream cards.

## Single-Step Benchmark Instruction

Apply a unified navy professional theme across the deck, make slide titles and card/body text stylistically consistent, and align the three workstream cards on Slide 2 onto one even horizontal line with equal spacing.

## Composite Atomic Operations

F4 Theme transfer, F1 Style consistency, F3 Layout control

## Composition Pattern

`retheme -> restyle -> realign`

## Atomic Scope

Only the deck-wide theming, style consistency, and Slide 2 card alignment are explicitly requested.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/pptc_001_retheme_restyle_and_realign_seed.pptx`.
- Keep the output format the same as the input document (`.pptx`).
- Focus on the requested composite operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .pptx deck with coherent theme, consistent styling, and corrected card alignment.
