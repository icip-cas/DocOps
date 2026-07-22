# Task

This task is adapted from the DocumentBenchmark seed `word_004`.

## Inputs

- Source document: `/root/word_004_transcript_style_consistency_seed.docx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This transcript packet was assembled from multiple contributors and now has visibly inconsistent formatting.

## Single-Step Benchmark Instruction

Make the document stylistically consistent by standardizing heading appearance and body-text formatting across all transcript sections, including paragraph-level indentation/spacing for the transcript body.

## Atomic Scope

Only style consistency is explicitly requested in this task.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/word_004_transcript_style_consistency_seed.docx`.
- Keep the output format the same as the input document (`.docx`).
- Focus on the requested atomic operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A cleaned-up .docx document with consistent run-level and paragraph-level styling.

## Latent Issues Intentionally Left In The Seed

- Capitalization of section labels is inconsistent
- Some spacing is awkward
- The content itself could be reorganized, but that is not the explicit request

## Inspired By

- IA for formatting? (https://www.reddit.com/r/word/comments/1rx137e/ia_for_formatting/)
