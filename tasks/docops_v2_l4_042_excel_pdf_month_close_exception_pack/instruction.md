# Task

This task is part of the DocumentBenchmark L4 manual expansion set.

## Goal

Prepare a release-ready monthly close workbook and matching PDF exception packet.

## Inputs

- `docops_v2_l4_042_excel_pdf_month_close_exception_pack_input.xlsx`: draft close workbook
- `receipt_manifest.xlsx`: authoritative receipt status source
- `invoice_evidence_packet.pdf`: source evidence pages
- `controller_close_review_note.docx`: authoritative controller overrides

## Required Outputs

Write both files to `/root/submission`:

- `docops_v2_l4_042_excel_pdf_month_close_exception_pack_input.xlsx`
- `month_close_exception_packet.pdf`

## Required Edit

Create `Clean Transactions` and `Exceptions` sheets. Classify receipt status from `receipt_manifest.xlsx` and controller notes:

- `RC-1002` is a duplicate exception for `$1,180`.
- `RC-1004` is a missing receipt exception for `$875`.
- `TRX-107` must be reclassified to `Consulting` and included as a `$1,975` policy exception.

The exception count is `3` and total exception amount is `$4,030`.

## Preservation Requirements

- Preserve live formulas in `Close Summary!B3:B8`.
- Use native Excel tables named `CleanTransactions` and `CloseExceptions`.
- Preserve data validation on `Clean Transactions!H2:H8`.
- Hide `Raw Transactions`, `Archive`, and `Rules`.
- Preserve the `receipt_status_choices` defined name.
- Highlight all review rows in `Clean Transactions`.
- The PDF packet must contain exactly the same three exceptions as the workbook and show the same count and total amount.
- Rebuild PDF bookmarks for cover and exception pages.
