# Task

This task is part of the DocumentBenchmark composite same-type split.

## Inputs

- Source document: `/root/excelc_001_ops_note_repair_and_evidence_highlight_seed.xlsx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This workbook contains an inaccurate ops-summary note that should be corrected and visually supported with evidence.

## Single-Step Benchmark Instruction

Using the evidence on `Monthly Data`, correct only cell `B8` on `Ops Summary` so the executive note reflects the true steepest backlog region and whether any site crossed the 12-incident threshold, then highlight only the supporting evidence rows on `Monthly Data`.

## Composite Atomic Operations

C5 Reasoning, C2 Editing, F2 Highlighting

## Composition Pattern

`understand -> edit -> highlight`

## Atomic Scope

Only the targeted note rewrite in `Ops Summary!B8` and the supporting-row highlight on `Monthly Data` are explicitly requested.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/excelc_001_ops_note_repair_and_evidence_highlight_seed.xlsx`.
- Keep the output format the same as the input document (`.xlsx`).
- Focus on the requested composite operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .xlsx workbook with an accurate executive-safe note and highlighted supporting rows.
