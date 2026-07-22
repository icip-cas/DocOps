# Task

This task is adapted from the DocumentBenchmark seed `ppt_003`.

## Inputs

- Source document: `/root/ppt_003_theme_transfer_seed.pptx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This deck was assembled from mismatched slides and needs a consistent visual theme.

## Single-Step Benchmark Instruction

Apply a unified blue professional theme across the whole presentation.

## Atomic Scope

Only global theme transfer is explicitly requested in this task.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/ppt_003_theme_transfer_seed.pptx`.
- Keep the output format the same as the input document (`.pptx`).
- Focus on the requested atomic operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .pptx deck with a coherent blue theme.

## Latent Issues Intentionally Left In The Seed

- Some spacing is awkward
- Slide layouts are basic
- Content wording could be improved, but that is not the explicit request

## Inspired By

- Why do so many presentations still have inconsistent colors and layouts? Are people not using Slide Master? (https://www.reddit.com/r/powerpoint/comments/1sknmk4/why_do_so_many_presentations_still_have/)
- Good all-round PowerPoint theme for lectures? (https://www.reddit.com/r/powerpoint/comments/1s7z1dz/good_allround_powerpoint_theme_for_lectures/)
