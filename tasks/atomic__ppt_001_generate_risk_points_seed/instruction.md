# Task

This task is adapted from the DocumentBenchmark seed `ppt_001`.

## Inputs

- Source document: `/root/ppt_001_generate_risk_points_seed.pptx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

The second slide is only a placeholder, but the first slide already contains enough context to draft operational risk points.

## Single-Step Benchmark Instruction

Generate three concise risk bullets on Slide 2 based on the information already present on Slide 1.

## Atomic Scope

Only content generation for Slide 2 is explicitly requested in this task.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/ppt_001_generate_risk_points_seed.pptx`.
- Keep the output format the same as the input document (`.pptx`).
- Focus on the requested atomic operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .pptx deck with generated risk bullets on Slide 2.

## Latent Issues Intentionally Left In The Seed

- The visual style is rough
- The boxes are not elegant
- The deck could use a proper theme, but that is not the explicit request

## Inspired By

- Need the best slides generator for work - any recommendations? (https://www.reddit.com/r/powerpoint/comments/1smn7nq/need_the_best_slides_generator_for_work_any/)
