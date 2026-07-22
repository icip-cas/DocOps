# Task

This task is adapted from the DocumentBenchmark seed `word_009`.

## Inputs

- Source document: `/root/word_009_clause_highlighting_seed.docx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This review note document contains several clauses, but only some are marked as critical.

## Single-Step Benchmark Instruction

Visually highlight the lines that are explicitly marked as Critical.

## Atomic Scope

Only emphasis/highlighting is explicitly requested in this task.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/word_009_clause_highlighting_seed.docx`.
- Keep the output format the same as the input document (`.docx`).
- Focus on the requested atomic operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .docx document with the correct lines highlighted.

## Why This Variant Is Hard

- The critical lines are embedded in otherwise plain text
- The task is local, so over-highlighting is a failure mode
- The document contains multiple clauses that look equally important but are not explicitly marked critical
