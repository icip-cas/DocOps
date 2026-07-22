# Task

This task is part of the DocumentBenchmark composite same-type split.

## Inputs

- Source document: `/root/wordc_001_repair_hierarchy_and_standardize_styles_seed.docx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This engineering report draft has both a broken heading hierarchy and inconsistent formatting.

## Single-Step Benchmark Instruction

Repair the heading hierarchy so the subsection under `2. Temporary bypass plan` restarts at `a. Isolation logic`, then standardize same-level heading styles and body paragraph formatting across the report.

## Composite Atomic Operations

S3 Hierarchy editing, F1 Style consistency

## Composition Pattern

`repair-structure -> restyle`

## Atomic Scope

Only the hierarchy repair and style consistency work are explicitly requested.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/wordc_001_repair_hierarchy_and_standardize_styles_seed.docx`.
- Keep the output format the same as the input document (`.docx`).
- Focus on the requested composite operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .docx document with corrected hierarchy and consistent heading/body styling.
