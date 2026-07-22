# Task

This task is adapted from the DocumentBenchmark seed `excel_008`.

## Inputs

- Source document: `/root/excel_008_client_update_editing_seed.xlsx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This workbook contains a draft client-facing status note that is too blunt for external communication.

## Single-Step Benchmark Instruction

Rewrite only cell `B8` on the `Status Draft` sheet so it sounds more formal and client-safe while preserving the underlying meaning about schedule uncertainty.

## Atomic Scope

Only local content editing is explicitly requested in this task.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/excel_008_client_update_editing_seed.xlsx`.
- Keep the output format the same as the input document (`.xlsx`).
- Focus on the requested atomic operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .xlsx workbook with the edited sentence in `B8`.

## Why This Variant Is Hard

- The edit must stay confined to one specific cell
- The rewritten sentence should soften tone without removing the risk signal
- Nearby cells include instructions and noise that should not be rewritten

## Inspired By

- is there a way to modify formulas based on inputs? (https://www.reddit.com/r/excel/comments/1snrbmh/is_there_a_way_to_modify_formulas_based_on_inputs/)
- Data Prep for Software Transition (https://www.reddit.com/r/excel/comments/1sna8on/data_prep_for_software_transition/)
