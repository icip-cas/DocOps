# Task

This task is part of the DocumentBenchmark composite same-type split.

## Inputs

- Source document: `/root/wordc_010_reason_and_rewrite_summary_note_seed.docx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This policy summary note misstates the true approval conflict and needs to be corrected.

## Single-Step Benchmark Instruction

Rewrite only the summary note paragraph so it accurately summarizes the approval conflict stated in the policy statements above, including the standard CFO rule for purchases above $50,000 and the emergency COO-only exception.

## Composite Atomic Operations

C5 Reasoning, C2 Editing

## Composition Pattern

`reason -> rewrite`

## Atomic Scope

Only the summary note rewrite grounded in the conflicting policy statements is explicitly requested.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/wordc_010_reason_and_rewrite_summary_note_seed.docx`.
- Keep the output format the same as the input document (`.docx`).
- Focus on the requested composite operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .docx policy memo with one corrected summary note paragraph.
