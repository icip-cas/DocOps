# Task

This task is adapted from the DocumentBenchmark seed `ppt_012`.

## Inputs

- Source document: `/root/ppt_012_outline_hierarchy_seed.pptx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This slide outline has the wrong bullet hierarchy: some items that should be sub-points are currently at the top level.

## Single-Step Benchmark Instruction

Fix the outline hierarchy so that task-level items are nested correctly under their parent workstreams.

## Atomic Scope

Only hierarchy editing is explicitly requested in this task.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/ppt_012_outline_hierarchy_seed.pptx`.
- Keep the output format the same as the input document (`.pptx`).
- Focus on the requested atomic operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .pptx deck with corrected bullet hierarchy.

## Why This Variant Is Hard

- The text content itself is fine, but the structural levels are wrong
- A visual fix is not enough if the logical bullet levels remain incorrect
- Multiple lines need consistent promotion/demotion decisions
