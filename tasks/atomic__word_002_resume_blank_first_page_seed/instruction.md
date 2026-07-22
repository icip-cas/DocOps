# Task

This task is adapted from the DocumentBenchmark seed `word_002`.

## Inputs

- Source document: `/root/word_002_resume_blank_first_page_seed.docx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This resume document opens with an unwanted blank first page before the actual content begins.

## Single-Step Benchmark Instruction

Delete the blank first page while keeping the rest of the resume content intact.

## Atomic Scope

Only deletion of the unwanted blank page is explicitly requested in this task.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/word_002_resume_blank_first_page_seed.docx`.
- Keep the output format the same as the input document (`.docx`).
- Focus on the requested atomic operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A corrected .docx document with no blank opening page.

## Latent Issues Intentionally Left In The Seed

- Fonts and section styling are inconsistent
- Spacing is slightly uneven across sections
- Footer page numbering may need review, but that is not the explicit request

## Inspired By

- How do you delete a blank first page?? (https://www.reddit.com/r/word/comments/1rhvhgd/how_do_you_delete_a_blank_first_page/)
- CV editing (https://www.reddit.com/r/word/comments/1s8t0is/cv_editing/)
