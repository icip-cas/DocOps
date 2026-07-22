# Task

This task is part of the DocumentBenchmark composite same-type split.

## Inputs

- Source document: `/root/pptc_002_reorder_storyline_and_add_exec_summary_seed.pptx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This deck should follow an executive story order and open with a constrained summary slide.

## Single-Step Benchmark Instruction

Reorder the deck to `Executive Summary -> Root Cause -> Implementation Plan -> Appendix`, then populate the `Executive Summary` slide with exactly three bullets using the prefixes `Vendor:`, `Inspection:`, and `Staffing:` in that order.

## Composite Atomic Operations

S2 Reorder, C3 Generation

## Composition Pattern

`reorder -> generate`

## Atomic Scope

Only the slide reorder and the constrained three-bullet summary content are explicitly requested.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/pptc_002_reorder_storyline_and_add_exec_summary_seed.pptx`.
- Keep the output format the same as the input document (`.pptx`).
- Focus on the requested composite operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .pptx deck with the required slide order and a grounded summary slide.
