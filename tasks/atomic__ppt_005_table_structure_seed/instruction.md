# Task

This task is adapted from the DocumentBenchmark seed `ppt_005`.

## Inputs

- Source document: `/root/ppt_005_table_structure_seed.pptx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This slide contains a milestone table, but the table is missing an `Owner` column.

## Single-Step Benchmark Instruction

Insert a new `Owner` column between `Milestone` and `Due Date` in the table on the slide.

## Atomic Scope

Only the table-structure change is explicitly requested in this task.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/ppt_005_table_structure_seed.pptx`.
- Keep the output format the same as the input document (`.pptx`).
- Focus on the requested atomic operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .pptx deck with the requested table modification.

## Why This Variant Is Hard

- Table edits in slides are brittle compared with plain text edits
- The change must preserve the existing rows and column order around the insertion point
- The slide is otherwise simple, so the task is specifically structural

## Inspired By

- Trying to figure out a PPT Report (https://www.reddit.com/r/powerpoint/comments/1smzxig/trying_to_figure_out_a_ppt_report/)
