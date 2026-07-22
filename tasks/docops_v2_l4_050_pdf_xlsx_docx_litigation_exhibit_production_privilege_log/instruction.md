# Task

Create the final litigation production package.

## Inputs

- `docops_v2_l4_050_pdf_xlsx_docx_litigation_exhibit_production_privilege_log_input.docx`: draft production memo with counsel-only notes
- `production_request_protocol.docx`: production protocol and Bates/privilege-log requirements
- `candidate_document_manifest.xlsx`: review-coded candidate list
- `source_candidate_documents.pdf`: unreviewed source bundle containing producible, privileged, duplicate, and non-responsive items
- `attorney_review_coding_notes.pdf`: final review-coding instructions

## Required Outputs

Write all three files to `/root/submission`:

- `produced_exhibit_binder.pdf`
- `production_privilege_log.xlsx`
- `production_index_memo.docx`

## Required Work

- Produce only `PROD-001`, `PROD-002`, `PROD-003`, and `PROD-004` in the PDF binder.
- Apply consecutive Bates labels `ACME-000001` through `ACME-000007`.
- Add PDF bookmarks at the first page of each produced document.
- Keep `CONFIDENTIAL` stamping on `PROD-001`, `PROD-003`, and `PROD-004`.
- Do not include `PRIV-005`, `PRIV-006`, `PRIV-007`, `NONRESP-008`, or `NONRESP-009` in the public PDF.
- The workbook must include a production index, privilege log, excluded-item sheet, QC summary formulas, data validations, conditional formatting, print areas, hidden rules, and hidden attorney notes.
- The DOCX memo must include a real TOC field, produced exhibit table, privilege/exclusion summary, highlighted QC paragraph, and cross-file reconciliation language.
