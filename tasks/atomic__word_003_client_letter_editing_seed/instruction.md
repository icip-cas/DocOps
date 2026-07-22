# Task

This task is adapted from the DocumentBenchmark seed `word_003`.

## Inputs

- Source document: `/root/word_003_client_letter_editing_seed.docx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This client communication draft contains rushed, overly blunt wording in the message body.

## Single-Step Benchmark Instruction

Rewrite the sections labeled `Draft opening` and `Draft ask` so they sound more professional and client-appropriate, while preserving the underlying facts and urgency.

## Atomic Scope

Only the text editing request is explicit in this task.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/word_003_client_letter_editing_seed.docx`.
- Keep the output format the same as the input document (`.docx`).
- Focus on the requested atomic operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .docx document with more professional wording in the specified sections.

## Latent Issues Intentionally Left In The Seed

- Fonts and spacing are inconsistent
- Heading formatting is uneven
- The document layout is visually messy

## Inspired By

- How much has Copilot changed the way word docs get edited these days? (https://www.reddit.com/r/word/comments/1skxkx3/how_much_has_copilot_changed_the_way_word_docs/)
