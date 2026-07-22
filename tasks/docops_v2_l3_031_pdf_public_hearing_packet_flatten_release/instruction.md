# Task

This task is part of a regenerated L3 single-document workflow set.

## Inputs

- Source PDF: `/root/docops_v2_l3_031_pdf_public_hearing_packet_flatten_release_input.pdf`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

Prepare a clean public hearing comment packet PDF from a messy internal fillable PDF.

## Required Sequential Workflow

This task must be completed in order:

1. Remove or flatten PDF form fields before publication.
2. Remove annotations, embedded files, JavaScript actions, and private PDF metadata.
3. Apply the correction memo before rebuilding the public pages.
4. Rebuild the corrected public hearing summary, comment category matrix, response workflow, and release controls.
5. Rebuild the PDF outline/bookmarks in the required page order.
6. Apply the required public visual style, headers, footers, and page numbering.
7. Save the clean public packet as a new PDF without modifying the source PDF.

## Single-PDF Workflow Instruction

Revise the source PDF into one clean public release packet:

- Save the revised PDF to `/root/submission/docops_v2_l3_031_pdf_public_hearing_packet_flatten_release_output.pdf`.
- Remove visible and hidden `DRAFT`, `PRIVATE`, `INTERNAL`, staff-only notes, personal phone numbers, private reviewer names, legal-hold material, and unreleased deliberation text.
- Remove AcroForm fields, PDF annotations, embedded files, JavaScript actions, and private metadata.
- Rebuild the PDF with exactly five pages in this order:
  1. `Public Hearing Packet: Zoning Amendment Comments`
  2. `Corrected Hearing Summary`
  3. `Comment Category Matrix`
  4. `Response Workflow`
  5. `Release Controls and Appendix Index`

## Required Corrections

Apply these corrections before rebuilding public pages:

- Hearing date must be `2026-03-14`, not `2026-03-04`.
- Public comment count must be `128`, not `112`.
- Response deadline must be `2026-09-15`, not `2026-09-01`.
- Public contact must be `Public Hearing Desk`, not individual staff names.
- Docket ID must be `ZA-2026-041`.

## Required PDF Features

- Rebuild bookmarks/outlines with the five required page titles.
- Use page size `Letter`.
- Use a public header bar color `0B3D91`.
- Use accent color `F59E0B` for public status callouts.
- Add footer text `Public Hearing Release | ZA-2026-041 | 2026` on every page.
- Set public PDF metadata:
  - Title: `Public Hearing Packet: Zoning Amendment Comments`
  - Author: `Public Hearing Desk`
  - Subject: `Clean flattened public release packet`

## Output Requirements

- Do not modify the source PDF in place.
- Keep the output format the same as the input PDF (`.pdf`).
- Complete the full ordered workflow, not just local text replacement.

## Expected Output Type

A revised `.pdf` artifact that completes a strict-order L3 single-document PDF release workflow with form-field flattening, annotation/attachment/JavaScript removal, correction application, public page rebuilding, bookmark rebuilding, metadata cleanup, and visual style migration.
