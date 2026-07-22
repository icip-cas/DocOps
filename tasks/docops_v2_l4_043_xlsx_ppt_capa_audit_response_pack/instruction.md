# Task

This task is part of the DocumentBenchmark L4 manual expansion set.

## Goal

Create a CAPA audit response tracker workbook and refresh the public audit response deck.

## Inputs

- `docops_v2_l4_043_xlsx_ppt_capa_audit_response_pack_input.pptx`: draft audit response PPTX with obsolete/internal slides
- `capa_deviation_log.xlsx`: authoritative deviation and CAPA source log
- `audit_scope_note.docx`: authoritative public-release boundary and owner alias mapping
- `regulation_excerpt.pdf`: authoritative clause wording
- `evidence_index.xlsx`: authoritative evidence readiness and public/internal status

## Required Outputs

Write both files to `/root/submission`:

- `capa_audit_tracker.xlsx`
- `docops_v2_l4_043_xlsx_ppt_capa_audit_response_pack_input.pptx`

## Required Edit

Create `capa_audit_tracker.xlsx` and refresh the deck so they agree:

- Public deck includes only CAPA-17, CAPA-18, and CAPA-24.
- CAPA-21/D-226/E-21A remain in the workbook but must not appear on public slides.
- Owner aliases must be expanded to canonical names.
- CAPA-17 is `At Risk`; CAPA-24 is `Blocked`; CAPA-18 is `Ready`; CAPA-21 is `Closed`.
- The public deck must remove the internal root-cause slide and keep the matrix template reference slide.

## Preservation Requirements

- `CAPA Tracker` must include native table `capa_audit_tracker` with ref `A1:K5`.
- `Evidence Readiness` must include native table `evidence_readiness` with ref `A1:F6`.
- Preserve formulas in `CAPA Tracker!J2:K5` and `Evidence Readiness!F2:F6`.
- Preserve data validation on `CAPA Tracker!G2:G5` and `H2:H5`.
- Preserve hidden `Rules` sheet plus defined names `public_status_choices` and `publish_choices`.
- Highlight CAPA-17 and CAPA-24 rows in the tracker.
- PPT and XLSX must agree on public CAPA IDs and must not leak internal root-cause details.
