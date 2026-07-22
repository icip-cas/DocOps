# Task

This task is part of the DocumentBenchmark composite same-type split.

## Inputs

- Source document: `/root/wordc_007_delete_blank_page_and_generate_profile_seed.docx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This resume opens with a blank page and lacks a concise professional profile.

## Single-Step Benchmark Instruction

Delete the blank first page, then insert exactly two profile lines under the candidate name: the first must start `Commissioning:` and the second must start `Vendor coordination:`.

## Composite Atomic Operations

S1 Insert/Delete, C3 Generation

## Composition Pattern

`delete -> generate`

## Atomic Scope

Only the blank-page deletion and the two-line profile generation are explicitly requested.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/wordc_007_delete_blank_page_and_generate_profile_seed.docx`.
- Keep the output format the same as the input document (`.docx`).
- Focus on the requested composite operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .docx resume with no blank opening page and a two-line profile under the name.
