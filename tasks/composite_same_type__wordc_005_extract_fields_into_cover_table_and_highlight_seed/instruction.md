# Task

This task is part of the DocumentBenchmark composite same-type split.

## Inputs

- Source document: `/root/wordc_005_extract_fields_into_cover_table_and_highlight_seed.docx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This shutdown memo needs the key fields surfaced at the top and the source evidence highlighted below.

## Single-Step Benchmark Instruction

Fill the cover table with the final approved shutdown date and final action owner, then highlight only the two evidence lines in the body that support those fields.

## Composite Atomic Operations

C1 Extraction, F2 Highlighting

## Composition Pattern

`extract -> highlight`

## Atomic Scope

Only the two extracted cover-table fields and the corresponding evidence highlights are explicitly requested.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/wordc_005_extract_fields_into_cover_table_and_highlight_seed.docx`.
- Keep the output format the same as the input document (`.docx`).
- Focus on the requested composite operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .docx memo with populated cover fields and highlighted evidence lines.
