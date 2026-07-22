# Task

This task is adapted from the DocumentBenchmark seed `pdf_004`.

## Inputs

- Source document: `/root/pdf_004_audit_pages_reorder_seed.pdf`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This audit PDF contains the right pages but in the wrong sequence.

## Single-Step Benchmark Instruction

Reorder the pages so the report flows in the sequence: Scope -> Findings -> Recommendations -> Appendix.

## Atomic Scope

Only page reordering is explicitly requested in this task.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/pdf_004_audit_pages_reorder_seed.pdf`.
- Keep the output format the same as the input document (`.pdf`).
- Focus on the requested atomic operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised PDF with the pages in correct order.

## Why This Variant Is Hard

- Every page looks valid on its own
- The correct order must be inferred from section titles rather than page numbers in the viewer
- No other content changes are requested

## Inspired By

- Slice and merge two different pdfs (https://www.reddit.com/r/pdf/comments/1s9p0hm/slice_and_merge_two_different_pdfs/)
