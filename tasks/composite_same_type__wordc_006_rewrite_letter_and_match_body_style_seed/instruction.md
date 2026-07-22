# Task

This task is part of the DocumentBenchmark composite same-type split.

## Inputs

- Source document: `/root/wordc_006_rewrite_letter_and_match_body_style_seed.docx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This client letter has two rough draft sections and inconsistent body formatting.

## Single-Step Benchmark Instruction

Rewrite the sections titled `Draft opening` and `Draft ask` so they sound professional and client-appropriate, then standardize those rewritten body paragraphs to one consistent style.

## Composite Atomic Operations

C2 Editing, F1 Style consistency

## Composition Pattern

`rewrite -> restyle`

## Atomic Scope

Only the two targeted rewrites and body-style normalization for those sections are explicitly requested.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/wordc_006_rewrite_letter_and_match_body_style_seed.docx`.
- Keep the output format the same as the input document (`.docx`).
- Focus on the requested composite operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .docx letter with polished draft sections and consistent body styling.
