# Task

This task is part of the DocumentBenchmark L4 manual expansion set.

## Goal

Finalize the external site inspection response letter and create the internal commitment tracker.

## Inputs

- `docops_v2_l4_067_word_xlsx_site_inspection_response_and_commitment_tracker_input.docx`: draft response letter
- `inspection_observation_log.xlsx`: authoritative inspection observation and draft commitment log
- `sponsor_comment_resolution_note.docx`: authoritative public/internal boundary and owner alias mapping
- `inspection_response_guidance.pdf`: response requirements
- `evidence_index_packet.pdf`: evidence IDs, release boundary, and status

## Required Outputs

Write both files to `/root/submission`:

- `docops_v2_l4_067_word_xlsx_site_inspection_response_and_commitment_tracker_input.docx`
- `site_commitment_tracker.xlsx`

## Required Edit

The external letter must include only OBS-01, OBS-02, and OBS-04. OBS-03 must remain only in the internal tracker because its evidence includes patient initials and randomization notes.

The letter must include a real TOC field, a response table, an evidence boundary section, and two highlighted urgent commitment paragraphs.

The Excel tracker must include all four commitments, expand owner aliases, preserve publish/internal decisions, and keep live formulas.

## Preservation Requirements

- `Commitment Tracker` must contain native table `site_commitment_tracker` with ref `A1:L5`.
- `Evidence Review` must contain native table `inspection_evidence_review` with ref `A1:F6`.
- Preserve formulas in `Commitment Tracker!K2:L5` and `Evidence Review!F2:F6`.
- Preserve data validation on `Commitment Tracker!H2:H5` and `J2:J5`.
- Preserve hidden `Rules` sheet plus defined names `status_choices` and `publish_choices`.
- Highlight tracker rows 2, 3, and 4.
- DOCX and XLSX must agree that only OBS-01, OBS-02, and OBS-04 are publishable in the letter.
