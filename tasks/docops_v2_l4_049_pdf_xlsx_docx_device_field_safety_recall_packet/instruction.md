# Task

Create the final field safety publication packet and internal reconciliation workbook.

## Inputs

- `docops_v2_l4_049_pdf_xlsx_docx_device_field_safety_recall_packet_input.docx`: draft customer field safety notice with stale and internal-only content
- `affected_lot_master.xlsx`: authoritative affected-product and publish/internal scope
- `customer_response_export.xlsx`: raw distributor/customer response export
- `field_action_regulatory_rules.pdf`: publication, privacy, and workbook-control rules
- `internal_root_cause_note.pdf`: internal-only root cause and legal-hold instructions

## Required Outputs

Write all three files to `/root/submission`:

- `customer_field_safety_notice.pdf`
- `field_action_reconciliation.xlsx`
- `regulatory_cover_memo.docx`

## Required Work

- The public PDF must publish only FA-01 through FA-06.
- The public PDF must not include FA-INT-07, FA-LEGAL-08, PILOT-77, LOT-2199X, ASIC comparator drift, attorney-work-product text, customer names, customer IDs, or open-unit counts.
- The public PDF must have four pages and PDF bookmarks for summary, affected products, customer actions, and response form/contact pages.
- The workbook must retain both internal rows, all customer response rows, formulas, native tables, print areas, data validation, conditional formatting, defined names, and hidden `Rules` plus `Internal Notes` sheets.
- The DOCX cover memo must include a real TOC field, public notice package table, highlighted high-risk follow-up paragraph, and cross-output reconciliation language.
- PDF, XLSX, and DOCX must agree that FA-01, FA-02, FA-03, FA-04, FA-05, and FA-06 are the only public notice rows.
