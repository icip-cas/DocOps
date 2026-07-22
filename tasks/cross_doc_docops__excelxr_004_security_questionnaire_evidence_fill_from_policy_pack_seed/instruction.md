# Task

This task is part of the DocumentBenchmark cross-document document-ops V7 realworld split.

## Goal

Complete the questionnaire by filling only the unanswered required controls from the supporting evidence.

## Primary Input

- `excelxr_004_security_questionnaire_evidence_fill_from_policy_pack_seed.xlsx`

## Supporting Inputs

- `security_questionnaire_extract.pdf`
- `followup_note.docx`
- `evidence_index.xlsx`

## Required Edit

Fill only the unanswered required rows in the questionnaire sheet using the response status, evidence ID, source page, and response note gathered from the supporting documents.

## Output

- Write the revised document to `/root/submission/excelxr_004_security_questionnaire_evidence_fill_from_policy_pack_seed.xlsx`
- Preserve non-target structure and formatting unless the task explicitly asks you to change it.

## Notes

- Leave prefilled non-required rows unchanged
- Keep the template sheet unchanged
- Do not replace summary formulas with hardcoded values
