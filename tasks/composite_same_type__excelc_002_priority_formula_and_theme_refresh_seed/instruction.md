# Task

This task is part of the DocumentBenchmark composite same-type split.

## Inputs

- Source document: `/root/excelc_002_priority_formula_and_theme_refresh_seed.xlsx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This tracker workbook needs a computed priority column and a coherent navy-and-steel visual treatment.

## Single-Step Benchmark Instruction

Insert a new `Priority Score` column between `Days Late` and `Owner` on `Tracker`, fill `D2:D5` with formulas that multiply the severity weight from `Weights` by `2` when `Days Late` is greater than `0` and otherwise by `1`, then apply a navy-and-steel theme to the header bands and visible sheet tabs.

## Composite Atomic Operations

S4 Table/Sheet ops, C4 Computation, F4 Theme transfer

## Composition Pattern

`restructure -> compute -> restyle`

## Atomic Scope

Only the inserted formula column and the requested workbook theme updates are explicitly requested.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/excelc_002_priority_formula_and_theme_refresh_seed.xlsx`.
- Keep the output format the same as the input document (`.xlsx`).
- Focus on the requested composite operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .xlsx workbook with the computed `Priority Score` column and consistent themed headers/tabs.
