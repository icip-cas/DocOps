# Task

This task is part of the DocumentBenchmark composite same-type split.

## Inputs

- Source document: `/root/pdfc_009_reason_and_rewrite_conflict_note_seed.pdf`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This contract note misstates the real payment conflict and needs one corrected summary sentence.

## Single-Step Benchmark Instruction

Rewrite only the cover note sentence so it accurately summarizes the payment-timing contradiction stated in the clause, schedule entry, and late-fee trigger below it.

## Composite Atomic Operations

C5 Reasoning, C2 Editing

## Composition Pattern

`reason -> rewrite`

## Atomic Scope

Only the single summary-note rewrite grounded in the conflicting payment terms is explicitly requested.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/pdfc_009_reason_and_rewrite_conflict_note_seed.pdf`.
- Keep the output format the same as the input document (`.pdf`).
- Focus on the requested composite operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .pdf contract note with one corrected conflict-summary sentence.
