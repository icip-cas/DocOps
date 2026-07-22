# Task

This task is part of the DocumentBenchmark composite same-type split.

## Inputs

- Source document: `/root/excelc_003_workbook_reorder_then_exec_summary_seed.xlsx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This workbook should follow an executive review order and include a tightly constrained summary sheet.

## Single-Step Benchmark Instruction

Reorder the workbook tabs to `Overview -> Detail_Q1 -> Detail_Q2 -> Risks`, then add a new sheet named `Executive Summary` at the end and fill `A2:A4` with exactly three short bullets using the prefixes `Supplier:`, `Inspection:`, and `Staffing:` in that order.

## Composite Atomic Operations

S2 Reorder, C3 Generation

## Composition Pattern

`reorder -> generate`

## Atomic Scope

Only the sheet reorder and the constrained three-bullet executive summary are explicitly requested.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/excelc_003_workbook_reorder_then_exec_summary_seed.xlsx`.
- Keep the output format the same as the input document (`.xlsx`).
- Focus on the requested composite operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .xlsx workbook with the required tab order and a grounded three-bullet summary sheet.
