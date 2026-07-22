# Task

Create the prior authorization appeal letter and appeal evidence tracker workbook.

## Inputs

- `docops_v2_l4_047_pdf_docx_xlsx_prior_auth_denial_appeal_packet_input.docx`
- `appeal_template.docx`
- `denial_letter.pdf`
- `payer_policy_excerpt.pdf`
- `neurology_visit_note.docx`
- `patient_impact_message.docx`
- `medication_and_claims.xlsx`
- `previous_authorization.pdf`
- `pulmonary_history_note.pdf`
- `internal_appeal_strategy.docx`

## Required Outputs

Write both files to `/root/submission`:

- `prior_auth_appeal_letter.docx`
- `appeal_evidence_tracker.xlsx`

## Required Work

- Appeal denial `DEN-7842` for Nina Patel, member `AC-77824`.
- Determine the correct 60-day appeal deadline from the denial date.
- Map the payer's migraine CGRP criteria to evidence from the clinical note, claims history, prior authorization history, and pulmonary safety note.
- Prove the request is for `Rimegepant ODT 75 mg`, `8 tablets per 30 days`, diagnosis `G43.709`.
- Build the workbook with native Excel tables, formulas, data validation, print areas, defined names, hidden raw/evidence rule sheets, and a privacy review.
- The external letter and public workbook sheets must not include internal legal strategy, DOI threat language, employer escalation language, or off-label workaround language.
