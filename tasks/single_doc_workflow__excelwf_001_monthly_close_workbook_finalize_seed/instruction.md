# Task

This task is part of the DocumentBenchmark single-document workflow split.

## User Request

Finalize a monthly close workbook.

## Primary Input

- `excelwf_001_monthly_close_workbook_finalize_seed.xlsx`

## Instructions

Add the helper formulas, reorder the tabs, normalize the issue log as a formal table named CloseIssues, add status dropdowns only to the live status cells, and leave a release-review comment on the summary sheet.

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

- Write the revised document to `/root/submission/excelwf_001_monthly_close_workbook_finalize_seed.xlsx`.
- Do not output a text-only answer; the required deliverable is the revised file.
