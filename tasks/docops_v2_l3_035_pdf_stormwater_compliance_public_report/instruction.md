# Task

This task is part of a regenerated L3 single-document workflow set.

## Inputs

- Source PDF: `/root/docops_v2_l3_035_pdf_stormwater_compliance_public_report_input.pdf`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

Prepare a clean public stormwater compliance report from a messy internal annual-review PDF.

## Single-PDF Workflow Instruction

Revise the source PDF into a polished public annual compliance report. Complete the full workflow in one output `.pdf`:

- Save the revised PDF to `/root/submission/docops_v2_l3_035_pdf_stormwater_compliance_public_report_output.pdf`.
- Remove all visible `DRAFT`, `PRIVATE`, `INTERNAL`, `scratch`, and confidential review material.
- Remove internal-only pages and rebuild the PDF with exactly these ten public pages in this order:
  1. `Riverbend Stormwater Compliance Report`
  2. `Executive Compliance Summary`
  3. `Permit Milestone Tracker`
  4. `Outfall Monitoring Results`
  5. `Corrective Action Register`
  6. `Public Complaint Log`
  7. `Green Infrastructure Inventory`
  8. `Sampling QA Checklist`
  9. `Publication Style Guide`
  10. `Appendix Release Register`
- Apply all correction-log fixes:
  `Outfal 07` to `Outfall 07`, `E. coli 410` to `E. coli 140`, corrective action due date `2026-09-01` to `2026-08-25`, green roof area `18,000 sq ft` to `18,800 sq ft`, and permit year `2025` to `2026`.

## Required Style Migration

The output must apply these explicit public-report style requirements:

- Use a 40-point full-width header band at the top of every page.
- Header band colors must be:
  - Cover and appendix pages: `243B53`
  - Summary and permit tracker pages: `007C89`
  - Monitoring, green infrastructure, and QA pages: `2E7D32`
  - Corrective action page: `C55A11`
  - Public complaint log page: `7030A0`
  - Publication style guide page: `5B6770`
- Put the page title in the header band.
- Use footer text `Riverbend Public Stormwater Report | Permit Year 2026` on every page.
- Use page numbering in the form `Page X of 10` on every page.
- Present public tables as clean pipe-delimited rows with corrected values.
- Set public PDF metadata title, subject, author, and keywords.

## Content Requirements

- Rebuild the milestone tracker, monitoring results, corrective action register, complaint log, green infrastructure inventory, QA checklist, style guide, and appendix register as public sections.
- Exclude enforcement strategy, contractor dispute, resident personal phone, legal settlement note, do-not-release notes, and unverified lab note text.

## Atomic Scope

Only the requested single-PDF workflow edits are in scope. Preserve public compliance facts, correction facts, page order, header-band style, footer text, page numbering, and PDF metadata requirements.

## Output Requirements

- Do not modify the source PDF in place.
- Keep the output format the same as the input PDF (`.pdf`).
- Complete the full PDF workflow, not just local text redaction.

## Expected Output Type

A revised `.pdf` artifact that completes a high-complexity L3 single-PDF compliance report rebuild, correction, cleanup, and explicit style migration workflow.
