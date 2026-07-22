# Task

This task is adapted from the DocumentBenchmark seed `pdf_008`.

## Inputs

- Source document: `/root/pdf_008_milestone_layout_control_seed.pdf`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This PDF contains three milestone callout boxes that are visually misaligned even though their content is correct.

## Single-Step Benchmark Instruction

Realign the three milestone boxes into one evenly spaced horizontal row beneath the title, without changing any of the box text.

## Atomic Scope

Only layout control is explicitly requested in this task.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/pdf_008_milestone_layout_control_seed.pdf`.
- Keep the output format the same as the input document (`.pdf`).
- Focus on the requested atomic operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised PDF with the milestone boxes cleanly aligned.

## Why This Variant Is Hard

- The page has a valid-looking but awkward layout rather than an obvious rendering failure
- The target requires geometric alignment, not text rewriting
- Text must remain unchanged while object positions move

## Inspired By

- Got stuck while building a PDF editing tool. (https://www.reddit.com/r/pdf/comments/1slyk16/got_stuck_while_building_a_pdf_editing_tool/)
