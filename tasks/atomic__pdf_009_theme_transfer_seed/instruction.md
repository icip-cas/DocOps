# Task

This task is adapted from the DocumentBenchmark seed `pdf_009`.

## Inputs

- Source document: `/root/pdf_009_theme_transfer_seed.pdf`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This PDF proposal uses inconsistent accent colors across pages and callout elements.

## Single-Step Benchmark Instruction

Apply a cohesive navy-based corporate theme across the whole PDF, including heading bars and callout boxes, while keeping the text content unchanged.

## Atomic Scope

Only theme transfer is explicitly requested in this task.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/pdf_009_theme_transfer_seed.pdf`.
- Keep the output format the same as the input document (`.pdf`).
- Focus on the requested atomic operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised PDF with a consistent document-wide theme.

## Why This Variant Is Hard

- The visual system spans multiple pages and multiple element types
- Local recoloring is not enough; the result should feel globally coherent
- Content and layout should stay intact while the theme changes

## Inspired By

- I need to translate a PDF file into another language while preserving the exact layout. (https://www.reddit.com/r/pdf/comments/1sm0hrg/i_need_to_translate_a_pdf_file_into_another/)
- I need Acrobat editor alternative (https://www.reddit.com/r/pdf/comments/1sja7un/i_need_acrobat_editor_alternative/)
