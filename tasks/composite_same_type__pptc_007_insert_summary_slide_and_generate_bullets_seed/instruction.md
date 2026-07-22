# Task

This task is part of the DocumentBenchmark composite same-type split.

## Inputs

- Source document: `/root/pptc_007_insert_summary_slide_and_generate_bullets_seed.pptx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This deck needs a new executive summary slide at the front.

## Single-Step Benchmark Instruction

Insert a new first slide titled `Executive Summary` and generate exactly three concise bullets using the prefixes `Supplier:`, `Inspection:`, and `Staffing:` in that order.

## Composite Atomic Operations

S1 Insert/Delete, C3 Generation

## Composition Pattern

`insert -> generate`

## Atomic Scope

Only the inserted first slide and its three generated bullets are explicitly requested.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/pptc_007_insert_summary_slide_and_generate_bullets_seed.pptx`.
- Keep the output format the same as the input document (`.pptx`).
- Focus on the requested composite operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .pptx deck with a new first summary slide containing exactly three bullets.
