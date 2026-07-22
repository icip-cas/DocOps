# Task

This task is adapted from the DocumentBenchmark seed `word_001`.

## Inputs

- Source document: `/root/word_001_engineering_report_toc_hierarchy_seed.docx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This engineering report draft has a broken heading hierarchy and a manually typed table of contents.

## Single-Step Benchmark Instruction

Repair the heading hierarchy so the document uses a consistent multi-level heading structure, the subsection under `2. Temporary bypass plan` restarts correctly from `a` instead of continuing from `c`, and the matching manually typed table-of-contents entry is updated to match.

## Atomic Scope

Only the heading hierarchy and its matching TOC entry are explicitly requested in this task.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/word_001_engineering_report_toc_hierarchy_seed.docx`.
- Keep the output format the same as the input document (`.docx`).
- Focus on the requested atomic operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A repaired .docx document with corrected heading hierarchy and matching TOC entry.

## Latent Issues Intentionally Left In The Seed

- The table of contents is manually typed
- Fonts and spacing are inconsistent
- Some headings use direct formatting instead of true styles

## Inspired By

- Custom Table of Content Template for my Engineering Report (https://www.reddit.com/r/word/comments/1s7j1d6/custom_table_of_content_template_for_my/)
- Headline issues (https://www.reddit.com/r/word/comments/1rzumyi/headline_issues/)
