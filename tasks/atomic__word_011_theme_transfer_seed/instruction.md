# Task

This task is adapted from the DocumentBenchmark seed `word_011`.

## Inputs

- Source document: `/root/word_011_theme_transfer_seed.docx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This document was assembled from mismatched source files and now looks visually inconsistent.

## Single-Step Benchmark Instruction

Apply a unified blue professional theme across the document.

## Atomic Scope

Only theme transfer is explicitly requested in this task.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/word_011_theme_transfer_seed.docx`.
- Keep the output format the same as the input document (`.docx`).
- Focus on the requested atomic operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .docx document with a coherent blue professional theme.

## Why This Variant Is Hard

- The document already has multiple conflicting style signals
- The requested output is global, not just a local formatting patch
- Theme transfer requires consistency without changing the underlying content
