# Task

This task is part of the DocumentBenchmark single-document workflow split.

## User Request

Clean up a capex register workbook.

## Primary Input

- `excelwf_004_capex_register_cleanup_seed.xlsx`

## Instructions

Insert the missing owner and status columns, compute the late-day formulas, convert the register into a formal table named CapexRegister, add status dropdowns to the live rows, and sort the register by risk so the workbook is publication ready.

Please complete the request directly in the input document. Treat the file as a working draft with explicit draft markers. Remove the `Draft:` prefixes, remove obsolete scratch material, repair the requested structure, and otherwise keep the live wording unchanged.

Use the visible sheet names and headers to restore the requested workbook structure; do not create extra summary sheets unless the task explicitly asks for one.
Put the sheets in the requested order and keep archive sheets hidden when they are part of the workbook.
Repair or add the needed formulas rather than hardcoding calculated outputs.
Highlight rows that need reviewer attention based on their status or exception meaning.
Make the active review sheets easier to navigate by freezing their header rows.
Convert the live review range into a formal Excel table with the named table specified in the request.
Add status dropdown validation only to the live status cells named in the task, not to entire columns.
Add reviewer-facing cell notes where the workbook asks for release comments.

## Output

- Write the revised document to `/root/submission/excelwf_004_capex_register_cleanup_seed.xlsx`.
- Do not output a text-only answer; the required deliverable is the revised file.
