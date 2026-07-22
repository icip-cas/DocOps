# Task

This task is part of the DocumentBenchmark single-document workflow split.

## User Request

Release the governance workbook without renaming or deleting its live sheets.

## Primary Input

- `excelwf_005_governance_workbook_release_seed.xlsx`

## Instructions

Restore the sheet order to Governance, Dashboard, and Archive; keep Archive hidden; repair the Dashboard formulas; convert the Governance range to a formal table named GovernanceMetrics; freeze the Dashboard header row; and add the release comment on the Dashboard title cell.

Please complete the request directly in the input document. Treat the file as a working draft with explicit draft markers. Remove the `Draft:` prefixes, remove obsolete scratch material, repair the requested structure, and otherwise keep the live wording unchanged.

Use the visible sheet names and headers to restore the requested workbook structure; do not create extra summary sheets unless the task explicitly asks for one.
Put the sheets in the requested order and keep archive sheets hidden when they are part of the workbook.
Repair or add the needed formulas rather than hardcoding calculated outputs.
Make the active review sheets easier to navigate by freezing their header rows.
Convert the live review range into a formal Excel table with the named table specified in the request.
Add reviewer-facing cell notes where the workbook asks for release comments.

## Output

- Write the revised document to `/root/submission/excelwf_005_governance_workbook_release_seed.xlsx`.
- Do not output a text-only answer; the required deliverable is the revised file.
