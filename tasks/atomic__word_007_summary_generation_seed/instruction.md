# Task

This task is adapted from the DocumentBenchmark seed `word_007`.

## Inputs

- Source document: `/root/word_007_summary_generation_seed.docx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This briefing document has context notes but the executive summary section is still only a placeholder.

## Single-Step Benchmark Instruction

Generate a concise executive summary paragraph based on the context notes already present in the document.

## Atomic Scope

Only content generation for the placeholder section is explicitly requested in this task.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/word_007_summary_generation_seed.docx`.
- Keep the output format the same as the input document (`.docx`).
- Focus on the requested atomic operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .docx document with a generated executive summary paragraph.

## Why This Variant Is Hard

- The source notes overlap and need synthesis rather than copying
- The generated paragraph must reflect both operational and commercial pressure
- The placeholder is explicit, but the right tone still requires judgment
