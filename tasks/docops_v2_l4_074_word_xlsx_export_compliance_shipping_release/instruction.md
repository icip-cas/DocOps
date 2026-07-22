# Task

Create the final export shipping release memo and export compliance release workbook.

## Inputs

- `docops_v2_l4_074_word_xlsx_export_compliance_shipping_release_input.docx`
- `commercial_invoice.xlsx`
- `packing_list.xlsx`
- `eccn_classification_note.pdf`
- `hts_origin_certificate.docx`
- `denied_party_screening.xlsx`
- `lithium_battery_sds.pdf`
- `freight_forwarder_email.docx`
- `end_use_statement.pdf`
- `internal_margin_strategy.pdf`

## Required Outputs

Write both files to `/root/submission`:

- `export_shipping_release_memo.docx`
- `export_compliance_release_workbook.xlsx`

## Required Work

- Release exactly three public lines: ELG-100, CRP-200, and BAT-310.
- Hold the prototype crypto board because it is 5A002 and requires license review.
- BAT-310 may release only with lithium handling documentation.
- Build a broker packet sequence that attaches only the seven external/public documents in order and excludes held/internal material.
- Add release reconciliation controls for released line count, held line count, internal-only count, released value, and BAT-310 lithium gate.
- Preserve workbook native tables, formulas, print areas, data validation, conditional formatting, hidden `Internal Pricing`, hidden `Rules`, and defined names.
- The memo must include the released-line table, broker packet assembly table, released value total of `$144,800.00`, and cross-file reconciliation statement.
- Do not expose internal margin, discount floor, target margin, reroute strategy, or license-risk appetite in public outputs.
