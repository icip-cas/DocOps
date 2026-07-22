# Task

This task is adapted from the DocumentBenchmark seed `ppt_010`.

## Inputs

- Source document: `/root/ppt_010_style_consistency_seed.pptx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This deck feels visually inconsistent from slide to slide.

## Single-Step Benchmark Instruction

Make the slide styling consistent across the deck, including titles, body text, and card appearance.

## Atomic Scope

Only style consistency is explicitly requested in this task.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/ppt_010_style_consistency_seed.pptx`.
- Keep the output format the same as the input document (`.pptx`).
- Focus on the requested atomic operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .pptx deck with consistent styling across slides.

## Why This Variant Is Hard

- Inconsistency appears across multiple style dimensions at once
- The deck is not completely broken, so the agent must normalize rather than redesign from scratch
- The request is global rather than local
