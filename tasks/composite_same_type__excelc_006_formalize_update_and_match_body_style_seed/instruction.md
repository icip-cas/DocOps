# Task

This task is part of the DocumentBenchmark composite same-type split.

## Inputs

- Source document: `/root/excelc_006_formalize_update_and_match_body_style_seed.xlsx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This client update sheet has one risky draft sentence and inconsistent nearby body formatting.

## Single-Step Benchmark Instruction

Rewrite only `Status Draft!B8` so it sounds formal and client-safe while preserving the schedule uncertainty, then make `B8:B10` use one consistent body style that matches `B7`.

## Composite Atomic Operations

C2 Editing, F1 Style consistency

## Composition Pattern

`rewrite -> restyle`

## Atomic Scope

Only the targeted draft rewrite and requested body-style normalization are explicitly requested.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/excelc_006_formalize_update_and_match_body_style_seed.xlsx`.
- Keep the output format the same as the input document (`.xlsx`).
- Focus on the requested composite operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .xlsx workbook with a polished sentence in `B8` and matched body styling in `B8:B10`.
