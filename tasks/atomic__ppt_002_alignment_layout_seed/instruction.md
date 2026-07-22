# Task

This task is adapted from the DocumentBenchmark seed `ppt_002`.

## Inputs

- Source document: `/root/ppt_002_alignment_layout_seed.pptx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This slide contains three content cards that are obviously misaligned and unevenly distributed.

## Single-Step Benchmark Instruction

Align the three workstream cards so they sit on a clean horizontal line with even spacing.

## Atomic Scope

Only layout/alignment is explicitly requested in this task.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/ppt_002_alignment_layout_seed.pptx`.
- Keep the output format the same as the input document (`.pptx`).
- Focus on the requested atomic operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .pptx deck with evenly aligned cards.

## Latent Issues Intentionally Left In The Seed

- Typography is plain
- Colors are inconsistent
- The deck has no master theme, but that is not the explicit request

## Inspired By

- Anyone else spending way too much time just formatting slides instead of actually thinking? (https://www.reddit.com/r/powerpoint/comments/1ryq4ko/anyone_else_spending_way_too_much_time_just/)
