# Task

This task is part of the DocumentBenchmark composite same-type split.

## Inputs

- Source document: `/root/pdfc_005_extract_vendor_fields_into_summary_panel_seed.pdf`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This vendor packet needs a concise extracted summary panel on the cover page.

## Single-Step Benchmark Instruction

Fill the cover `Key Facts` panel with vendor name, quote reference, total including options, and delivery window extracted from the detailed quote page.

## Composite Atomic Operations

C1 Extraction, C3 Generation

## Composition Pattern

`extract -> generate`

## Atomic Scope

Only the four extracted fields inserted into the cover summary panel are explicitly requested.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/pdfc_005_extract_vendor_fields_into_summary_panel_seed.pdf`.
- Keep the output format the same as the input document (`.pdf`).
- Focus on the requested composite operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .pdf packet with a populated cover `Key Facts` panel.
