# Task

This task is adapted from the DocumentBenchmark seed `word_010`.

## Inputs

- Source document: `/root/word_010_image_layout_seed.docx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This document contains three site images, but they are not laid out cleanly.

## Single-Step Benchmark Instruction

Adjust the layout so the three images appear on one neat row with balanced spacing and alignment.

## Atomic Scope

Only layout control is explicitly requested in this task.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/word_010_image_layout_seed.docx`.
- Keep the output format the same as the input document (`.docx`).
- Focus on the requested atomic operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .docx document with a cleaner image layout.

## Why This Variant Is Hard

- The content itself is simple, but the visual arrangement is awkward
- The images have uneven sizing and alignment
- A local layout fix can easily disturb surrounding structure if handled clumsily
