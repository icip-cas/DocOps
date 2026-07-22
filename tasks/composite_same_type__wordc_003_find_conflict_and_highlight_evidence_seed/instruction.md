# Task

This task is part of the DocumentBenchmark composite same-type split.

## Inputs

- Source document: `/root/wordc_003_find_conflict_and_highlight_evidence_seed.docx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This policy note contains an approval-logic conflict that should be highlighted for legal review.

## Single-Step Benchmark Instruction

Highlight only the two statements that conflict on whether emergency purchases above $50,000 need CFO approval and leave the rest of the wording unchanged.

## Composite Atomic Operations

C5 Reasoning, F2 Highlighting

## Composition Pattern

`detect -> highlight`

## Atomic Scope

Only the true conflict-evidence highlight is explicitly requested.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/wordc_003_find_conflict_and_highlight_evidence_seed.docx`.
- Keep the output format the same as the input document (`.docx`).
- Focus on the requested composite operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .docx document with the conflicting approval statements highlighted.
