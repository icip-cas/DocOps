# Task

This task is part of the DocumentBenchmark composite same-type split.

## Inputs

- Source document: `/root/pdfc_006_rewrite_paragraph_and_match_body_style_seed.pdf`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This status memo contains one informal paragraph and inconsistent body styling.

## Single-Step Benchmark Instruction

Rewrite only the paragraph beginning `The supplier missed two promised dates again` so it reads formally while preserving the meaning, then make its body style match the surrounding body paragraphs.

## Composite Atomic Operations

C2 Editing, F1 Style consistency

## Composition Pattern

`rewrite -> restyle`

## Atomic Scope

Only the targeted paragraph rewrite and body-style matching are explicitly requested.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/pdfc_006_rewrite_paragraph_and_match_body_style_seed.pdf`.
- Keep the output format the same as the input document (`.pdf`).
- Focus on the requested composite operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .pdf memo with one formalized paragraph and matching body styling.
