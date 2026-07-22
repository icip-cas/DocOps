# Task

This task is adapted from the DocumentBenchmark seed `excel_011`.

## Inputs

- Source document: `/root/excel_011_theme_transfer_seed.xlsx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This workbook has three operational tabs with inconsistent accent colors and no shared visual theme.

## Single-Step Benchmark Instruction

Apply a muted navy-and-steel visual theme across all visible sheets in this workbook, including section headers and tab colors, while keeping the workbook content unchanged.

## Atomic Scope

Only global theme transfer is explicitly requested in this task.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/excel_011_theme_transfer_seed.xlsx`.
- Keep the output format the same as the input document (`.xlsx`).
- Focus on the requested atomic operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .xlsx workbook with a consistent cross-sheet theme.

## Why This Variant Is Hard

- The workbook spans multiple sheets with different existing accent colors
- A good result needs global visual coherence, not one-off recoloring
- Content and structure should remain intact while the visual system changes

## Inspired By

- Search Bar thats acts like a Slicer (https://www.reddit.com/r/excel/comments/1snq9ob/search_bar_thats_acts_like_a_slicer/)
