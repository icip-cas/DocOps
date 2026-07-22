# Task

This task is part of a regenerated L3 single-document workflow set.

## Inputs

- Source document: `/root/docops_v2_l3_003_docx_public_records_release_metadata_cleanup_input.docx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

Prepare a clean public records release packet from a messy internal Word document about water main break reimbursement claims.

## Required Sequential Workflow

This task must be completed in order:

1. Accept or resolve tracked changes before rebuilding the public packet.
2. Delete all comments and comment references.
3. Remove private metadata from document properties.
4. Apply the correction memo before rebuilding the public facts, tables, workflow, and appendix index.
5. Remove all PII, legal-strategy notes, draft-only content, and private staff names.
6. Rebuild the public release packet with the required section order, tables, footer, page breaks, and styles.
7. Save the clean public document as a new `.docx` without modifying the source document.

## Single-Document Workflow Instruction

Revise the source document into one clean public release packet:

- Save the revised document to `/root/submission/docops_v2_l3_003_docx_public_records_release_metadata_cleanup_output.docx`.
- Remove visible and hidden `DRAFT`, `PRIVATE`, `INTERNAL`, tracked-change remnants, comments, legal-strategy notes, personal phone numbers, home addresses, claimant names, and internal staff names.
- Remove private document properties and set public metadata only.
- Rebuild the document with exactly these public headings in this order:
  1. `Public Release Packet: Water Main Break Claims`
  2. `Release Summary`
  3. `Corrected Determination Facts`
  4. `Claim Category Table`
  5. `Notification Workflow`
  6. `Redaction and Metadata Checklist`
  7. `Public Appendix Index`

## Required Corrections

Apply these corrections before rebuilding public content:

- Incident date must be `2026-01-18`, not `2026-01-08`.
- Claim count must be `37`, not `34`.
- Reimbursable cap must be `$650`, not `$500`.
- Public contact must be `Public Records Desk`, not individual staff names.
- Response deadline must be `2026-08-15`, not `2026-07-31`.

## Required Public Tables

Rebuild three public tables:

- `Corrected Determination Facts`
- `Claim Category Table`
- `Redaction and Metadata Checklist`

The tables must use the corrected public facts and must not contain private notes or old draft values.

## Required Style and Layout

- Use `Aptos Display` bold title text, 18 pt, color `1F4E79`.
- Use `Aptos` body text, 10.5 pt.
- Use heading text in color `1F4E79`.
- Use table header fill `1F4E79` with white bold text.
- Set all page margins to `0.75` inches.
- Add footer text `Public Records Release | Water Main Break Claims | 2026`.
- Insert page breaks before `Redaction and Metadata Checklist` and `Public Appendix Index`.

## Required Metadata Cleanup

- Remove Word comments and comment relationships.
- Remove tracked change XML (`w:ins`, `w:del`, and comment range tags).
- Set document author to `Public Records Office`.
- Set document subject to `Clean public release packet`.
- Do not leave private keywords, staff names, draft notes, or legal-hold metadata in document properties.

## Output Requirements

- Do not modify the source document in place.
- Keep the output format the same as the input document (`.docx`).
- Complete the full ordered workflow, not just visible text replacement.

## Expected Output Type

A revised `.docx` artifact that completes a strict-order L3 single-document public records release workflow with tracked-change cleanup, comment removal, metadata cleanup, correction application, table rebuilding, footer insertion, and style migration.
