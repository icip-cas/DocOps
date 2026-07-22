# Task

This task is adapted from the DocumentBenchmark seed `ppt_006`.

## Inputs

- Source document: `/root/ppt_006_delete_outdated_slide_seed.pptx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This deck still contains one outdated appendix slide from a previous reporting cycle.

## Single-Step Benchmark Instruction

Delete the slide titled `Old Appendix (DELETE ME)`.

## Atomic Scope

Only slide deletion is explicitly requested in this task.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/ppt_006_delete_outdated_slide_seed.pptx`.
- Keep the output format the same as the input document (`.pptx`).
- Focus on the requested atomic operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .pptx deck with the outdated slide removed.

## Why This Variant Is Hard

- The slide is fully formatted, so deletion requires confidence rather than obvious corruption
- The rest of the deck should remain unchanged
- Page order must stay coherent after deletion

## Inspired By

- Live Event Production: Managing PowerPoint decks (https://www.reddit.com/r/powerpoint/comments/1rzzm9l/live_event_production_managing_powerpoint_decks/)
