# Task

This task is part of the DocumentBenchmark L4 manual expansion set.

## Goal

Create a final policy exception release memo and a public PDF appendix.

## Inputs

- `docops_v2_l4_064_word_pdf_policy_exception_memo_public_appendix_input.docx`: draft policy exception memo
- `policy_exception_log.xlsx`: authoritative exception log
- `policy_source_excerpt.pdf`: policy control source
- `approval_public_release_note.docx`: authoritative approval and public-release boundary
- `evidence_appendix_draft.pdf`: evidence appendix containing both public and internal-only content

## Required Outputs

Write both files to `/root/submission`:

- `docops_v2_l4_064_word_pdf_policy_exception_memo_public_appendix_input.docx`
- `public_policy_exception_appendix.pdf`

## Required Edit

Create a final memo with Scope, Release Matrix, Exception Register, and Appendix A. Replace the manual TOC placeholder with a real TOC field. Header must include `Open blocking exceptions: 2`.

The public PDF appendix must publish only EX-102, EX-118, and EX-131. It must exclude EX-144 and all internal legal strategy or compensating-control details.

## Preservation Requirements

- Preserve DOCX heading structure and section/page break before Appendix A.
- Highlight the two blocking exception paragraphs.
- Memo Exception Register may include closed EX-144 for internal tracking.
- Public PDF appendix must not include closed/internal-only material.
- DOCX and PDF must agree on the public exception IDs.
- Rebuild PDF bookmarks for the appendix.
