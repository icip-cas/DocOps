# Task

This task is part of the DocumentBenchmark composite same-type split.

## Inputs

- Source document: `/root/pdfc_001_payment_conflict_highlight_seed.pdf`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This contract excerpt contains a payment-timing conflict that should be highlighted for review.

## Single-Step Benchmark Instruction

Review the payment references on page 1 and highlight only the statements that create the timing contradiction between the base clause, the schedule entry, and the late-fee trigger. Do not rewrite the text.

## Composite Atomic Operations

C5 Reasoning, F2 Highlighting

## Composition Pattern

`detect -> highlight`

## Atomic Scope

Only the true conflicting payment lines should be visually highlighted.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/pdfc_001_payment_conflict_highlight_seed.pdf`.
- Keep the output format the same as the input document (`.pdf`).
- Focus on the requested composite operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .pdf document with the conflicting payment lines highlighted.
