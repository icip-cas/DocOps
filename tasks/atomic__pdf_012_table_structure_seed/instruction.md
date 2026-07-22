# Task

This task is adapted from the DocumentBenchmark seed `pdf_012`.

## Inputs

- Source document: `/root/pdf_012_table_structure_seed.pdf`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This PDF contains a structured table where one column incorrectly combines two logical fields.

## Single-Step Benchmark Instruction

Split the `Owner / Team` column into two separate table columns named `Owner` and `Team`, while preserving the existing row content.

## Atomic Scope

Only table-structure editing is explicitly requested in this task.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/pdf_012_table_structure_seed.pdf`.
- Keep the output format the same as the input document (`.pdf`).
- Focus on the requested atomic operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised PDF with the table split into the requested columns.

## Why This Variant Is Hard

- The PDF table is visually valid even though its schema is wrong
- Each row must be restructured without losing any cell content
- This is a structural table operation, not a free-form rewrite

## Inspired By

- how can i transfer a pdf table into word? (https://www.reddit.com/r/word/comments/1si8saz/how_can_i_transfer_a_pdf_table_into_word/)
- Tool to extract images from PDFs in their original format, without additional compression, and while retaining color profile information? (https://www.reddit.com/r/pdf/comments/1smqv12/tool_to_extract_images_from_pdfs_in_their/)
