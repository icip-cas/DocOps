# Task

This task is adapted from the DocumentBenchmark seed `ppt_004`.

## Inputs

- Source document: `/root/ppt_004_slide_reorder_seed.pptx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

The deck's slides are out of logical order.

## Single-Step Benchmark Instruction

Reorder the slides so the presentation flows in the sequence: Executive Summary -> Root Cause -> Implementation Plan -> Appendix.

## Atomic Scope

Only slide reordering is explicitly requested in this task.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/ppt_004_slide_reorder_seed.pptx`.
- Keep the output format the same as the input document (`.pptx`).
- Focus on the requested atomic operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .pptx deck with corrected slide order.

## Latent Issues Intentionally Left In The Seed

- The deck has no polished theme
- The slide layouts are very plain
- Some text could be tightened, but that is not the explicit request

## Inspired By

- Automated table of contents (https://www.reddit.com/r/powerpoint/comments/1s98sok/automated_table_of_contents/)
