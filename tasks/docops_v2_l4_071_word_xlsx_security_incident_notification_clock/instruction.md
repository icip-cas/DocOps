# Task

Create the final security incident notification memo and breach notification tracker.

## Inputs

- `docops_v2_l4_071_word_xlsx_security_incident_notification_clock_input.docx`: draft privileged incident memo with unapproved language
- `incident_timeline.xlsx`: event timeline and evidence IDs
- `security_notice_terms.pdf`: contract, processor, law-enforcement, and public-communication rules
- `forensic_summary.docx`: confirmed technical facts and affected data types
- `privileged_incident_strategy_note.pdf`: internal-only legal and exploit detail boundaries

## Required Outputs

Write both files to `/root/submission`:

- `incident_notification_memo.docx`
- `breach_notification_tracker.xlsx`

## Required Work

- Use discovery timestamp `2026-06-12 09:15`.
- Track five notification work items and their due times.
- Publicly state only the confirmed affected data: names, email addresses, and order IDs for 1,248 customer records.
- State that payment card data, Social Security numbers, and password hashes were not confirmed exposed.
- Preserve Excel native tables, formulas, print areas, data validation, conditional formatting, hidden `Privileged Strategy`, hidden `Rules`, and defined names.
- The DOCX must include a real TOC, notification clock table, highlighted Acme Retail deadline paragraph, approved public-language section, and cross-file reconciliation.
- Do not expose exploit payload, CVE label, zero-day wording, root password detail, legal strategy, or EV-IR-03/EV-IR-04 in public materials.
