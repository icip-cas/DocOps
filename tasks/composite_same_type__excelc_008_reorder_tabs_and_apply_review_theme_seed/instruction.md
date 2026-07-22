# Task

This task is part of the DocumentBenchmark composite same-type split.

## Inputs

- Source document: `/root/excelc_008_reorder_tabs_and_apply_review_theme_seed.xlsx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This review pack should follow a clean tab order and use a cohesive blue theme.

## Single-Step Benchmark Instruction

Reorder the workbook tabs to `Cover -> Metrics -> Risks -> Notes`, then apply a muted blue theme to the sheet title bands and visible tab colors.

## Composite Atomic Operations

S2 Reorder, F4 Theme transfer

## Composition Pattern

`reorder -> retheme`

## Atomic Scope

Only the requested sheet reorder and visual theme transfer are explicitly requested.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/excelc_008_reorder_tabs_and_apply_review_theme_seed.xlsx`.
- Keep the output format the same as the input document (`.xlsx`).
- Focus on the requested composite operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .xlsx workbook with the target tab order and a blue review theme.
