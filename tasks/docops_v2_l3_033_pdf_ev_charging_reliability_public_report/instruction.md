# Task

This task is part of a regenerated L3 single-document workflow set.

## Inputs

- Source PDF: `/root/docops_v2_l3_033_pdf_ev_charging_reliability_public_report_input.pdf`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

Prepare a clean public EV charging reliability report from a messy internal network operations PDF.

## Required Sequential Workflow

This task must be completed in order. Later sections depend on corrected facts from earlier steps:

1. Remove all internal-only pages and all confidential text.
2. Apply the correction-log facts before rebuilding any public section.
3. Rebuild the reliability summary from corrected uptime and station counts.
4. Rebuild the station uptime register using the corrected station ID and uptime.
5. Rebuild the outage timeline using the corrected payment outage date.
6. Rebuild the maintenance action log using the corrected technician owner.
7. Rebuild equity access, payment incident, communications, and appendix sections from the corrected public facts.
8. Apply the required public-report style and PDF metadata after the public content is rebuilt.

## Single-PDF Workflow Instruction

Revise the source PDF into a polished public reliability report. Complete the full workflow in one output `.pdf`:

- Save the revised PDF to `/root/submission/docops_v2_l3_033_pdf_ev_charging_reliability_public_report_output.pdf`.
- Remove all visible `DRAFT`, `PRIVATE`, `INTERNAL`, `scratch`, and confidential operations material.
- Remove internal-only pages and rebuild the PDF with exactly these ten public pages in this order:
  1. `MetroCharge EV Reliability Report`
  2. `Reliability Summary`
  3. `Station Uptime Register`
  4. `Outage Event Timeline`
  5. `Maintenance Action Log`
  6. `Equity Access Commitments`
  7. `Payment System Incident Review`
  8. `Public Communications Timeline`
  9. `Publication Style Guide`
  10. `Appendix Release Register`
- Apply all correction-log fixes before rebuilding the public pages:
  `Station C-17` to `Station C-71`, uptime `91.2%` to `97.2%`, payment outage date `2026-02-18` to `2026-02-08`, technician `Nia Rowe` to `Nia Rowan`, and public notice date `2026-03-12` to `2026-03-10`.

## Required Style Migration

The output must apply these explicit public-report style requirements:

- Use a 40-point full-width header band at the top of every page.
- Header band colors must be:
  - Cover and appendix pages: `0B132B`
  - Reliability summary and station uptime pages: `0F766E`
  - Outage and payment incident pages: `B45309`
  - Maintenance page: `1D4ED8`
  - Equity and communications pages: `6D28D9`
  - Publication style guide page: `4B5563`
- Put the page title in the header band.
- Use footer text `MetroCharge Public Reliability Report | Q1 2026` on every page.
- Use page numbering in the form `Page X of 10` on every page.
- Present public tables as clean pipe-delimited rows with corrected values.
- Set public PDF metadata title, subject, author, and keywords.

## Content Requirements

- Rebuild the reliability summary, station uptime register, outage timeline, maintenance action log, equity access commitments, payment system incident review, public communications timeline, publication style guide, and appendix register as public sections.
- Exclude vendor penalty note, driver complaint phone, cybersecurity exception, legal hold, do-not-release language, internal SLA dispute, and private station revenue text.

## Output Requirements

- Do not modify the source PDF in place.
- Keep the output format the same as the input PDF (`.pdf`).
- Complete the full ordered workflow, not just local text redaction.

## Expected Output Type

A revised `.pdf` artifact that completes a strict-order L3 single-PDF EV charging reliability report workflow with correction, rebuilding, metadata cleanup, and explicit style migration.
