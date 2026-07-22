# Task

This task is adapted from the DocumentBenchmark seed `word_005`.

## Inputs

- Source document: `/root/word_005_meeting_minutes_extraction_seed.docx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This meeting-minutes draft contains several proposed dates and ownership notes, but some were superseded later in the document.

## Single-Step Benchmark Instruction

Extract the final approved shutdown date and the final action owner for the vendor contact list.

## Atomic Scope

Only information extraction is explicitly requested in this task.

## Output Requirements

- Do not modify the source document in place.
- Write your final answer to `/root/submission/final_answer.txt`.
- Keep the answer concise, directly responsive, and grounded in the provided document.

## Expected Output Type

A concise extraction of the two requested fields.

## Why This Variant Is Hard

- Earlier values are plausible but outdated
- Corrections appear later and are phrased narratively, not in a clean table
- There are near-miss distractors like tentative and fallback dates

## Inspired By

- How much has Copilot changed the way word docs get edited these days? (https://www.reddit.com/r/word/comments/1skxkx3/how_much_has_copilot_changed_the_way_word_docs/)
