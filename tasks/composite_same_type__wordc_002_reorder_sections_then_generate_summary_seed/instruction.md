# Task

This task is part of the DocumentBenchmark composite same-type split.

## Inputs

- Source document: `/root/wordc_002_reorder_sections_then_generate_summary_seed.docx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This report is out of order and needs a constrained executive summary at the front.

## Single-Step Benchmark Instruction

Reorder the document so it reads `Executive Summary -> Findings -> Next Steps -> Appendix Materials`, then place exactly three short summary lines directly under `Executive Summary` using the prefixes `Permit:`, `Supplier:`, and `Safety:` in that order.

## Composite Atomic Operations

S2 Reorder, C3 Generation

## Composition Pattern

`reorder -> generate`

## Atomic Scope

Only the section reorder and one-paragraph executive summary generation are explicitly requested.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/wordc_002_reorder_sections_then_generate_summary_seed.docx`.
- Keep the output format the same as the input document (`.docx`).
- Focus on the requested composite operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .docx document with the requested section order and one grounded summary paragraph.
