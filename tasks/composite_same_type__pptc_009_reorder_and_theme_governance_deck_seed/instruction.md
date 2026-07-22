# Task

This task is part of the DocumentBenchmark composite same-type split.

## Inputs

- Source document: `/root/pptc_009_reorder_and_theme_governance_deck_seed.pptx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This governance deck should follow a cleaner story order and use one navy theme.

## Single-Step Benchmark Instruction

Reorder the slides to `Overview -> Risks -> Actions -> Appendix`, then apply a unified navy theme by turning the title text boxes into navy title bars and the main body text boxes into matching light-blue cards across the deck.

## Composite Atomic Operations

S2 Reorder, F4 Theme transfer

## Composition Pattern

`reorder -> retheme`

## Atomic Scope

Only the requested slide reorder and navy theme transfer are explicitly requested.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/pptc_009_reorder_and_theme_governance_deck_seed.pptx`.
- Keep the output format the same as the input document (`.pptx`).
- Focus on the requested composite operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .pptx deck with corrected order and a navy visual theme.
