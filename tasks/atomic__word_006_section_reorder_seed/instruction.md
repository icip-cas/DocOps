# Task

This task is adapted from the DocumentBenchmark seed `word_006`.

## Inputs

- Source document: `/root/word_006_section_reorder_seed.docx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This report draft has major sections in the wrong order.

## Single-Step Benchmark Instruction

Reorder the document sections so they appear in the logical sequence: Executive Summary -> Findings -> Next Steps -> Appendix Materials.

## Atomic Scope

Only section reordering is explicitly requested in this task.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/word_006_section_reorder_seed.docx`.
- Keep the output format the same as the input document (`.docx`).
- Focus on the requested atomic operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .docx document with corrected section order.

## Why This Variant Is Hard

- Every section already looks complete, so the problem is structural rather than missing content
- The current numbering is misleading because the physical order does not match the intended narrative flow
- Reordering must preserve the section content itself

## Inspired By

- Sort lists in Word without needing Excel (https://www.reddit.com/r/word/comments/1qf4pat/sort_lists_in_word_without_needing_excel/)
