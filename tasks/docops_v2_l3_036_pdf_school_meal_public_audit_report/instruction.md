# Task

This task is part of a regenerated L3 single-document workflow set.

## Inputs

- Source PDF: `/root/docops_v2_l3_036_pdf_school_meal_public_audit_report_input.pdf`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

Prepare a clean public school meal program audit report from a messy internal review PDF.

## Single-PDF Workflow Instruction

Revise the source PDF into a polished public audit report. Complete the full workflow in one output `.pdf`:

- Save the revised PDF to `/root/submission/docops_v2_l3_036_pdf_school_meal_public_audit_report_output.pdf`.
- Remove all visible `DRAFT`, `PRIVATE`, `INTERNAL`, `scratch`, and confidential review material.
- Remove internal-only pages and rebuild the PDF with exactly these eleven public pages in this order:
  1. `Metro Schools Meal Program Public Audit`
  2. `Executive Findings`
  3. `Meal Count Reconciliation`
  4. `Site Compliance Matrix`
  5. `Vendor Invoice Review`
  6. `Allergen Communication Log`
  7. `Corrective Action Timeline`
  8. `Public Outreach Commitments`
  9. `Nutrition Standard Checklist`
  10. `Publication Style Guide`
  11. `Appendix Release Register`
- Apply all correction-log fixes:
  meal count `18,420` to `18,240`, `Site B-12` to `Site B-21`, invoice overage `$9,800` to `$980`, allergen notice date `2026-05-03` to `2026-05-13`, and corrective due date `2026-06-30` to `2026-06-15`.

## Required Style Migration

The output must apply these explicit public-report style requirements:

- Use a 40-point full-width header band at the top of every page.
- Header band colors must be:
  - Cover and appendix pages: `1F2937`
  - Executive findings and meal reconciliation pages: `0F766E`
  - Site compliance and nutrition checklist pages: `166534`
  - Vendor invoice page: `7C2D12`
  - Allergen communication and public outreach pages: `1D4ED8`
  - Corrective action page: `B91C1C`
  - Publication style guide page: `6B7280`
- Put the page title in the header band.
- Use footer text `Metro Schools Public Meal Audit | School Year 2026` on every page.
- Use page numbering in the form `Page X of 11` on every page.
- Present public tables as clean pipe-delimited rows with corrected values.
- Set public PDF metadata title, subject, author, and keywords.

## Content Requirements

- Rebuild the findings, meal count reconciliation, site compliance, invoice review, allergen log, corrective action, outreach commitments, nutrition checklist, style guide, and appendix register as public sections.
- Exclude student medical note, vendor dispute, staff disciplinary note, personal phone, legal settlement draft, do-not-release language, and internal scoring memo.

## Atomic Scope

Only the requested single-PDF workflow edits are in scope. Preserve public audit facts, correction facts, page order, header-band style, footer text, page numbering, and PDF metadata requirements.

## Output Requirements

- Do not modify the source PDF in place.
- Keep the output format the same as the input PDF (`.pdf`).
- Complete the full PDF workflow, not just local text redaction.

## Expected Output Type

A revised `.pdf` artifact that completes a high-complexity L3 single-PDF public audit report rebuild, correction, cleanup, and explicit style migration workflow.
