# Task

This task is adapted from the DocumentBenchmark seed `pdf_006`.

## Inputs

- Source document: `/root/pdf_006_management_summary_generation_seed.pdf`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This PDF contains raw operational findings plus an empty management summary area.

## Single-Step Benchmark Instruction

Replace the placeholder inside the blank `Management Summary` box on page 1 with a concise executive summary.

The summary must mention all three themes:
- recurring file-quality issues, including the outdated mapping/header inconsistency
- approval delays caused by an outdated approver list
- inconsistent ticket labeling across queues

## Atomic Scope

Only content generation is explicitly requested in this task.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/pdf_006_management_summary_generation_seed.pdf`.
- Keep the output format the same as the input document (`.pdf`).
- Focus on the requested atomic operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised PDF with a constrained executive summary added to page 1.

## Why This Variant Is Hard

- The target summary is not directly present anywhere in the PDF
- Relevant evidence is distributed across multiple raw notes
- The output must fit a constrained, pre-existing summary space

## Inspired By

- Export Reddit thread as a PDF? (https://www.reddit.com/r/pdf/comments/1slym6z/export_reddit_thread_as_a_pdf/)
- Tools for Academic Literature PDFs to Plain Text? (https://www.reddit.com/r/pdf/comments/1sjyvhv/tools_for_academic_literature_pdfs_to_plain_text/)
