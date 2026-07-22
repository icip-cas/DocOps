# Task

This task is part of the DocumentBenchmark composite same-type split.

## Inputs

- Source document: `/root/excelc_004_outline_and_print_ready_program_plan_seed.xlsx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This program plan should become collapsible and print cleanly on a single landscape page.

## Single-Step Benchmark Instruction

On `Program Plan`, make rows `4`, `7`, and `10` the top-level section rows, group the following detail rows under each section using row outlining, then set the sheet to print on one landscape page with row `2` repeated.

## Composite Atomic Operations

S3 Hierarchy editing, F3 Layout control

## Composition Pattern

`restructure -> print-layout`

## Atomic Scope

Only the row-outline hierarchy and print-layout setup are explicitly requested.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/excelc_004_outline_and_print_ready_program_plan_seed.xlsx`.
- Keep the output format the same as the input document (`.xlsx`).
- Focus on the requested composite operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .xlsx workbook with outline grouping and a print-ready page setup.
