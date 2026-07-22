# Task

This task is adapted from the DocumentBenchmark seed `pdf_011`.

## Inputs

- Source document: `/root/pdf_011_outline_hierarchy_seed.pdf`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This PDF already has bookmarks, but the bookmark hierarchy does not match the section structure of the document.

## Single-Step Benchmark Instruction

Update the PDF bookmarks so `Findings` and `Recommendations` are top-level entries, and `Site A` and `Site B` appear nested under `Findings`.

## Atomic Scope

Only hierarchy editing in the PDF outline / bookmark structure is explicitly requested in this task.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/pdf_011_outline_hierarchy_seed.pdf`.
- Keep the output format the same as the input document (`.pdf`).
- Focus on the requested atomic operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised PDF with corrected bookmark hierarchy.

## Why This Variant Is Hard

- The content pages are correct, but the structural navigation layer is wrong
- The incorrect outline is not obvious unless the PDF bookmark tree is inspected
- A correct fix changes hierarchy without rewriting the visible document text

## Inspired By

- Looking for Best Way to Attach Files to PDF Document (https://www.reddit.com/r/pdf/comments/1sl4n2w/looking_for_best_way_to_attach_files_to_pdf/)
- Preview pdf fast or in large grid view? (https://www.reddit.com/r/pdf/comments/1sijife/preview_pdf_fast_or_in_large_grid_view/)
