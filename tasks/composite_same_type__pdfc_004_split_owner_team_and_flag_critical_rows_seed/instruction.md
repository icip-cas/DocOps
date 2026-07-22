# Task

This task is part of the DocumentBenchmark composite same-type split.

## Inputs

- Source document: `/root/pdfc_004_split_owner_team_and_flag_critical_rows_seed.pdf`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This action table needs separate owner/team columns and stronger emphasis on the critical rows.

## Single-Step Benchmark Instruction

Split the `Owner / Team` header into separate `Owner` and `Team` columns, preserve the row wording, and visually emphasize only the two rows marked `Critical`.

## Composite Atomic Operations

S4 Table/Sheet ops, F2 Highlighting

## Composition Pattern

`restructure -> emphasize`

## Atomic Scope

Only the column split and emphasis of the critical rows are explicitly requested.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/pdfc_004_split_owner_team_and_flag_critical_rows_seed.pdf`.
- Keep the output format the same as the input document (`.pdf`).
- Focus on the requested composite operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .pdf document with separate `Owner` and `Team` columns and highlighted critical rows.
