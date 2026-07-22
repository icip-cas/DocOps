# Task

This task is adapted from the DocumentBenchmark seed `ppt_007`.

## Inputs

- Source document: `/root/ppt_007_cross_slide_extraction_seed.pptx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This deck contains multiple mentions of dates and owners, but some values were later corrected.

## Single-Step Benchmark Instruction

Extract the final launch date and the final action owner from the presentation.

## Atomic Scope

Only information extraction is explicitly requested in this task.

## Output Requirements

- Do not modify the source document in place.
- Write your final answer to `/root/submission/final_answer.txt`.
- Keep the answer concise, directly responsive, and grounded in the provided document.

## Expected Output Type

A concise extraction of the requested fields.

## Why This Variant Is Hard

- Early slides contain plausible but outdated values
- The correct answer is distributed across multiple slides
- The deck requires cross-slide resolution rather than local reading
