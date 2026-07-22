# Task

This task is part of the DocumentBenchmark L4 manual expansion set.

## Goal

Prepare the invoice packet PDF for client release using the release register, redaction policy, and approved appendix.

## Primary Input

- `docops_v2_l4_055_pdf_invoice_packet_redaction_from_vendor_register_input.pdf`

## Supporting Inputs

- `vendor_release_register.xlsx`
- `client_redaction_policy.docx`
- `approved_remittance_appendix.pdf`

## Required Edit

Remove the duplicate Alpine Labs invoice page and the internal payment routing page. Keep the approved Alpine Labs and Birch Clinic invoice pages.

Redact Federal Tax ID and Bank Routing values on released invoice pages using textual redaction markers:

- `[REDACTED TAX ID]`
- `[REDACTED BANK ROUTING]`

Insert `approved_remittance_appendix.pdf` after the released invoice pages. Update the cover to `Status: Final Client Release` and `Page Count: 4`.

## Output

- Write the revised PDF to `/root/submission/docops_v2_l4_055_pdf_invoice_packet_redaction_from_vendor_register_output.pdf`
- Preserve non-target invoice text and rebuild bookmarks for the final packet.

## Notes

- The final page order must be cover, Alpine invoice, Birch invoice, approved remittance appendix.
- Do not include duplicate, draft-only, or internal-only pages.
- Do not leave raw tax IDs or bank routing numbers visible.
