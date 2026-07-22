# Task

This task is part of a regenerated L3 single-document workflow set.

## Inputs

- Source PDF: `/root/docops_v2_l3_034_pdf_permit_notice_packet_errata_repair_input.pdf`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

Prepare a clean public permit notice repair packet from a messy draft PDF with OCR-like spelling errors, wrong facts, duplicate pages, and obsolete values.

## Required Sequential Workflow

1. Remove duplicate draft pages and private scratch material.
2. Apply the errata register before rebuilding any public page.
3. Correct spelling, parcel, applicant, date, fee, inspection, and condition errors throughout the packet.
4. Rebuild the six public pages in the required order.
5. Rebuild public outline bookmarks.
6. Apply the required header/footer style and public metadata.
7. Save the clean PDF as a new file without modifying the source PDF.

## Single-PDF Workflow Instruction

Save the revised PDF to `/root/submission/docops_v2_l3_034_pdf_permit_notice_packet_errata_repair_output.pdf`.

Remove all visible `DRAFT`, `PRIVATE`, `scratch`, duplicate-page, OCR-error, obsolete, and do-not-post material. Rebuild exactly these six public pages:

1. `Public Permit Notice Repair Packet`
2. `Corrected Errata Register`
3. `Applicant and Parcel Summary`
4. `Inspection Schedule`
5. `Fee and Condition Matrix`
6. `Public Posting Checklist`

## Required Corrections

- `Greenvile` must become `Greenville`.
- `permti`, `publc`, `adress`, `inspeciton`, `reciept`, and `conditon` must be corrected.
- Applicant must be `Arbor House Cooperative`.
- Parcel ID must be `P-2046-17B`, not `P-2046-17D`.
- Public address must be `412 Cedar Avenue`.
- Notice posting date must be `2026-08-14`, not `2026-08-04`.
- Hearing date must be `2026-09-18`, not `2026-09-08`.
- Inspection date must be `2026-08-29`, not `2026-08-19`.
- Filing fee must be `$1,850`, not `$1,580`.
- Public condition count must be `7`, not `5`.

## Required PDF Features

- Page size must be Letter.
- Use header bar color `234E70` and accent color `F59E0B`.
- Add footer text `Public Permit Notice Repair Packet | 2026` on every page.
- Rebuild the outline/bookmarks with the six page titles.
- Set PDF metadata title to `Public Permit Notice Repair Packet`, author to `Permit Records Office`, and subject to `Clean errata repair packet`.

## Output Requirements

- Do not modify the source PDF in place.
- Keep the output format `.pdf`.
- Complete the full ordered workflow, not just local spelling replacement.
