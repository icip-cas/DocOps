# Task

This task is adapted from the DocumentBenchmark seed `ppt_011`.

## Inputs

- Source document: `/root/ppt_011_highlight_key_item_seed.pptx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This milestone slide contains one explicitly critical item that should stand out more clearly.

## Single-Step Benchmark Instruction

Visually emphasize the line marked as Critical so it stands out from the other bullets.

## Atomic Scope

Only highlighting/emphasis is explicitly requested in this task.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/ppt_011_highlight_key_item_seed.pptx`.
- Keep the output format the same as the input document (`.pptx`).
- Focus on the requested atomic operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .pptx deck with the critical line clearly highlighted.

## Why This Variant Is Hard

- The task is narrow and local, so over-formatting is an easy failure mode
- The slide already has multiple bullets competing for attention
- The agent must identify the single intended emphasis target
