# Task

This task is adapted from the DocumentBenchmark seed `pdf_005`.

## Inputs

- Source document: `/root/pdf_005_paragraph_editing_seed.pdf`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This PDF contains a client-facing draft where one paragraph is too blunt for external use.

## Single-Step Benchmark Instruction

Rewrite only the paragraph that begins with `The supplier missed two promised dates again`.

The replacement paragraph must:
- keep these facts: two promised dates were missed, the latest package was incomplete, and the timeline remains uncertain until revised materials are confirmed
- use a more formal client-safe tone
- avoid emotionally charged wording such as `again` and `cannot rely on the timeline they gave us`

## Atomic Scope

Only local content editing is explicitly requested in this task.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/pdf_005_paragraph_editing_seed.pdf`.
- Keep the output format the same as the input document (`.pdf`).
- Focus on the requested atomic operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised PDF with the targeted paragraph rewritten using constrained factual anchors.

## Why This Variant Is Hard

- The edit is scoped to one paragraph inside a fixed-layout PDF
- The tone must soften without removing the schedule risk signal
- Other imperfections in the PDF are intentionally left in place and are not the explicit request

## Inspired By

- How do you edit a pdf document while keeping everything the same font and layout exactly? (https://www.reddit.com/r/pdf/comments/1snrzd5/how_do_you_edit_a_pdf_document_while_keeping/)
- I want to keep the formatting (https://www.reddit.com/r/pdf/comments/1sg500e/i_want_to_keep_the_formatting/)
