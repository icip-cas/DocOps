# Task

This task is adapted from the DocumentBenchmark seed `pdf_010`.

## Inputs

- Source document: `/root/pdf_010_appendix_delete_seed.pdf`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This PDF packet contains one obsolete appendix page that should not be part of the final document.

## Single-Step Benchmark Instruction

Delete the page titled `Superseded Rate Card` and leave the remaining pages unchanged.

## Atomic Scope

Only page deletion is explicitly requested in this task.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/pdf_010_appendix_delete_seed.pdf`.
- Keep the output format the same as the input document (`.pdf`).
- Focus on the requested atomic operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised PDF with the obsolete page removed.

## Why This Variant Is Hard

- The obsolete page still looks like a legitimate report page
- The operation is structural, not textual
- The surrounding pages should remain in place and unchanged

## Inspired By

- Slice and merge two different pdfs (https://www.reddit.com/r/pdf/comments/1s9p0hm/slice_and_merge_two_different_pdfs/)
- Acrobat Pro alternatives for batch Bates Numbering across multiple PDFs? (https://www.reddit.com/r/pdf/comments/1sm47um/acrobat_pro_alternatives_for_batch_bates/)
