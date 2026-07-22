# Task

This task is adapted from the DocumentBenchmark seed `pdf_007`.

## Inputs

- Source document: `/root/pdf_007_body_style_consistency_seed.pdf`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This PDF brief has consistent section labels but inconsistent body-paragraph styling.

## Single-Step Benchmark Instruction

Make the body paragraphs on page 1 use one consistent body style, while leaving the title and section labels unchanged.

## Atomic Scope

Only formatting consistency is explicitly requested in this task.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/pdf_007_body_style_consistency_seed.pdf`.
- Keep the output format the same as the input document (`.pdf`).
- Focus on the requested atomic operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised PDF with normalized body-paragraph styling.

## Why This Variant Is Hard

- The styling drift looks partly intentional because different paragraphs use different emphasis systems
- The instruction targets body text only, not every text element on the page
- Visual consistency matters more than textual changes here

## Inspired By

- White text isn't white. Have to use pptx workaround to make it white? (https://www.reddit.com/r/pdf/comments/1sm0tyd/white_text_isnt_white_have_to_use_pptx_workaround/)
- Wobbly text when set at an angle slight angle using indesign (https://www.reddit.com/r/pdf/comments/1sm0nby/wobbly_text_when_set_at_an_angle_slight_angle/)
