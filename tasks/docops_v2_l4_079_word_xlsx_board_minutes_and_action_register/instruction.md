# Task

This task is part of the DocumentBenchmark L4 manual expansion set.

## Goal

Publish final board minutes and a matching action register workbook.

## Inputs

- `docops_v2_l4_079_word_xlsx_board_minutes_and_action_register_input.docx`: draft minutes document
- `published_agenda_packet.pdf`: authoritative section order
- `attendance_vote_action_log.xlsx`: authoritative attendance, vote, and action seed rows
- `chair_reconciliation_notes.docx`: authoritative conflict overrides

## Required Outputs

Write both files to `/root/submission`:

- `docops_v2_l4_079_word_xlsx_board_minutes_and_action_register_input.docx`
- `board_action_register.xlsx`

## Required Edit

Create final minutes in this section order:

1. Publication Status
2. Attendance
3. Motions
4. Action Register
5. Appendix A - Posting Package

Use the chair notes to resolve conflicts:

- M-17 was approved as amended, not pending.
- M-18 was deferred, not approved.
- Priya Shah attended remotely and must not be shown as absent.
- A-18 must remain `Watch`.

## Preservation Requirements

- Replace the manual TOC placeholder with a real TOC field.
- Set the footer status to `Minutes packet status: Final`.
- Start Appendix A on a new section/page.
- Highlight the Watch-item paragraph for A-18.
- The Word Action Register table and Excel `Action Register` rows must match exactly.
- The Excel output must contain a native table named `board_action_register`.
- Preserve formulas in `Action Register!G2:G4`.
- Preserve data validation on `Action Register!E2:E4`.
- Preserve the hidden `Rules` sheet and `status_choices` defined name.
