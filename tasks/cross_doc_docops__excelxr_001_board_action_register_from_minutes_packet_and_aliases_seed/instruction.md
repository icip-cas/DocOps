# Task

This task is part of the DocumentBenchmark cross-document document-ops V7 realworld split.

## Goal

Populate only the blank rows of the `Decision Register` sheet with motion ID.

## Primary Input

- `excelxr_001_board_action_register_from_minutes_packet_and_aliases_seed.xlsx`

## Supporting Inputs

- `board_packet_extract.pdf`
- `decision_note.docx`
- `owner_map.xlsx`

## Required Edit

Populate only the blank rows of the `Decision Register` sheet with motion ID, agenda item, normalized owner, due date, and packet-support section from the meeting materials. Then move `Decision Register` directly after `Cover` without altering formulas, styles, or other sheets.

## Output

- Write the revised document to `/root/submission/excelxr_001_board_action_register_from_minutes_packet_and_aliases_seed.xlsx`
- Preserve non-target structure and formatting unless the task explicitly asks you to change it.

## Notes

- Keep the template sheets unchanged
- Do not replace formulas with hardcoded text
