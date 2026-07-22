# Task

This task is adapted from the DocumentBenchmark seed `pdf_001`.

## Inputs

- Source document: `/root/pdf_001_vendor_quote_extraction_seed.pdf`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This vendor quote PDF contains key commercial details in a messy, semi-structured layout.

## Single-Step Benchmark Instruction

Extract the requested fields from the PDF and reply using exactly this 6-line template:

`Vendor: ...`
`Quote Ref: ...`
`Quoted Total Including Options: ...`
`Delivery Window: ...`
`Warranty Term: ...`
`Project Contact Email: ...`

## Atomic Scope

Only information extraction is explicitly requested in this task.

## Output Requirements

- Do not modify the source document in place.
- Write your final answer to `/root/submission/final_answer.txt`.
- Keep the answer concise, directly responsive, and grounded in the provided document.

## Expected Output Type

A structured, template-constrained extraction of the requested fields.

## Latent Issues Intentionally Left In The Seed

- The document layout is visually messy
- Relevant fields are split across narrative text and tables
- There are extra commercial notes that are easy to confuse with the requested output

## Inspired By

- Extracting data from PDF in an organized manner? (https://www.reddit.com/r/excel/comments/1sn21r5/extracting_data_from_pdf_in_an_organized_manner/)
- how can i transfer a pdf table into word? (https://www.reddit.com/r/word/comments/1si8saz/how_can_i_transfer_a_pdf_table_into_word/)
