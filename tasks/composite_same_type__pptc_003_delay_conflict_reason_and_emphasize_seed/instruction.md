# Task

This task is part of the DocumentBenchmark composite same-type split.

## Inputs

- Source document: `/root/pptc_003_delay_conflict_reason_and_emphasize_seed.pptx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This deck contains one outdated delay statement that should be visually emphasized for review.

## Single-Step Benchmark Instruction

Compare the updated launch date on Slide 1 with the schedule bullets on Slide 2, then visually emphasize only the inconsistent bullet on Slide 2 so it stands out from the surrounding lines.

## Composite Atomic Operations

C5 Reasoning, F2 Highlighting

## Composition Pattern

`detect -> emphasize`

## Atomic Scope

Only the emphasis of the truly inconsistent delay line is explicitly requested.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/pptc_003_delay_conflict_reason_and_emphasize_seed.pptx`.
- Keep the output format the same as the input document (`.pptx`).
- Focus on the requested composite operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .pptx deck with the inconsistent line clearly emphasized.
