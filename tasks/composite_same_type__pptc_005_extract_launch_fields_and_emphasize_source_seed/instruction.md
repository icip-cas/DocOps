# Task

This task is part of the DocumentBenchmark composite same-type split.

## Inputs

- Source document: `/root/pptc_005_extract_launch_fields_and_emphasize_source_seed.pptx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This deck needs the final launch date and owner surfaced on the summary slide and the source line emphasized.

## Single-Step Benchmark Instruction

Fill the two summary callouts on Slide 1 with the final launch date and final action owner from the deck, then visually emphasize the source line on the detailed status slide.

## Composite Atomic Operations

C1 Extraction, F2 Highlighting

## Composition Pattern

`extract -> emphasize`

## Atomic Scope

Only the two extracted callout values and the corresponding source-line emphasis are explicitly requested.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/pptc_005_extract_launch_fields_and_emphasize_source_seed.pptx`.
- Keep the output format the same as the input document (`.pptx`).
- Focus on the requested composite operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .pptx deck with populated summary callouts and one emphasized evidence line.
