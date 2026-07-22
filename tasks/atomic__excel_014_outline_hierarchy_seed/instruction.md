# Task

This task is adapted from the DocumentBenchmark seed `excel_014`.

## Inputs

- Source document: `/root/excel_014_outline_hierarchy_seed.xlsx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This workbook contains a flat program outline that has section rows and detail rows, but no usable hierarchy in Excel's outline structure.

## Single-Step Benchmark Instruction

On the `Program Outline` sheet, make rows `4`, `7`, and `10` the top-level section rows and group the following detail rows under each section using Excel row outlining, without changing the row order.

## Atomic Scope

Only hierarchy editing is explicitly requested in this task.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/excel_014_outline_hierarchy_seed.xlsx`.
- Keep the output format the same as the input document (`.xlsx`).
- Focus on the requested atomic operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .xlsx workbook with row grouping / outline hierarchy applied.

## Why This Variant Is Hard

- The hierarchy is visually implied but not structurally encoded
- The task requires identifying the correct parent-child row boundaries
- A correct result changes outline structure rather than rewriting content

## Inspired By

- Making recipe tree chart but having trouble formatting and automating it (https://www.reddit.com/r/excel/comments/1snlk2n/making_recipe_tree_chart_but_having_trouble/)
