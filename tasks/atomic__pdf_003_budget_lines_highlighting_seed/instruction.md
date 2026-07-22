# Task

This task is adapted from the DocumentBenchmark seed `pdf_003`.

## Inputs

- Source document: `/root/pdf_003_budget_lines_highlighting_seed.pdf`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This approval sheet contains several spend lines, but only some should be visually emphasized for review.

## Single-Step Benchmark Instruction

Highlight the budget lines whose amounts are strictly greater than $10,000.

## Atomic Scope

Only visual emphasis/highlighting is explicitly requested in this task.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/pdf_003_budget_lines_highlighting_seed.pdf`.
- Keep the output format the same as the input document (`.pdf`).
- Focus on the requested atomic operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised PDF with the correct lines highlighted.

## Why This Variant Is Hard

- Near-threshold distractors are present, including $9,900 and exactly $10,000
- The rule is strict inequality, not greater-than-or-equal
- The content is simple, so the difficulty comes from precise interpretation rather than document length

## Inspired By

- How can you remove watermarks from a PDF? (https://www.reddit.com/r/pdf/comments/1smiay5/how_can_you_remove_watermarks_from_a_pdf/)
