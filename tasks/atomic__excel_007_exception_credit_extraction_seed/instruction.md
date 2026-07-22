# Task

This task is adapted from the DocumentBenchmark seed `excel_007`.

## Inputs

- Source document: `/root/excel_007_exception_credit_extraction_seed.xlsx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This workbook contains an exception tracker plus an unstructured email trail with superseded values.

## Single-Step Benchmark Instruction

Extract the final approved customer credit amount for request `CR-1184` if the North Unit shipment goes out after Apr 18.

Reply using exactly this 3-line template:

`Request ID: CR-1184`
`Condition: shipment goes out after Apr 18`
`Final approved credit: $6,950`

## Atomic Scope

Only content extraction is explicitly requested in this task.

## Output Requirements

- Do not modify the source document in place.
- Write your final answer to `/root/submission/final_answer.txt`.
- Keep the answer concise, directly responsive, and grounded in the provided document.

## Expected Output Type

A concise, template-constrained extraction.

## Why This Variant Is Hard

- The workbook contains a near-match request ID (`CR-1184A`) that is not the target
- An earlier draft amount is still visible in the tracker but has been superseded in the notes
- The target value depends on a condition embedded inside a narrative email block

## Inspired By

- Extracting data from PDF in an organized manner? (https://www.reddit.com/r/excel/comments/1sn21r5/extracting_data_from_pdf_in_an_organized_manner/)
- Is there a way to search for values across sheets? (https://www.reddit.com/r/excel/comments/1smwxwu/is_there_a_way_to_search_for_values_across_sheets/)
