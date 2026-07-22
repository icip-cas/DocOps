# Task

This task is part of the DocumentBenchmark composite same-type split.

## Inputs

- Source document: `/root/excelc_005_extract_credit_amount_and_highlight_source_seed.xlsx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This credit workbook needs the final approved amount surfaced and the source decision row highlighted.

## Single-Step Benchmark Instruction

Fill `Summary!B3` with the final approved customer credit amount for `CR-1184` when the shipment date is after `2026-04-18`, then highlight the supporting decision row on `Requests`.

## Composite Atomic Operations

C1 Extraction, F2 Highlighting

## Composition Pattern

`extract -> highlight`

## Atomic Scope

Only the targeted extracted amount and supporting-row highlight are explicitly requested.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/excelc_005_extract_credit_amount_and_highlight_source_seed.xlsx`.
- Keep the output format the same as the input document (`.xlsx`).
- Focus on the requested composite operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .xlsx workbook with the extracted amount in `Summary!B3` and the correct source row highlighted.
