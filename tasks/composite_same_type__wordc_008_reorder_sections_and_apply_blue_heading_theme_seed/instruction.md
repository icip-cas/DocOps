# Task

This task is part of the DocumentBenchmark composite same-type split.

## Inputs

- Source document: `/root/wordc_008_reorder_sections_and_apply_blue_heading_theme_seed.docx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This report needs both a better section sequence and a consistent blue heading theme.

## Single-Step Benchmark Instruction

Reorder the sections to `Executive Summary -> Findings -> Next Steps -> Appendix`, then apply a blue heading theme to all top-level section headings.

## Composite Atomic Operations

S2 Reorder, F4 Theme transfer

## Composition Pattern

`reorder -> retheme`

## Atomic Scope

Only the requested section reorder and top-level heading theme change are explicitly requested.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/wordc_008_reorder_sections_and_apply_blue_heading_theme_seed.docx`.
- Keep the output format the same as the input document (`.docx`).
- Focus on the requested composite operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .docx report with corrected section order and blue top-level headings.
