# Task

Create the final carton artwork proof and QA checklist.

## Inputs

- `docops_v2_l4_051_pdf_xlsx_carton_artwork_locale_proof_input.pdf`: draft PDF with wrong page order and retired-market artwork
- `sku_panel_master.xlsx`: authoritative SKU/panel list
- `locale_typography_note.docx`: locale and publication rules
- `carton_artwork_style_sheet.pdf`: proof layout and bookmark requirements

## Required Outputs

Write both files to `/root/submission`:

- `docops_v2_l4_051_pdf_xlsx_carton_artwork_locale_proof_input.pdf`
- `carton_artwork_qa_checklist.xlsx`

## Required Style/Format Work

- Final PDF must contain cover, SKU-A12 English, SKU-A12 Spanish, SKU-B07 English, and QA Certificate pages.
- Exclude SKU-X99 / Legacy French Panel from final PDF.
- Rebuild PDF bookmarks.
- Preserve Spanish panel warning status in the Excel checklist.
- Excel must include native table, formulas, print area, freeze panes, data validation, conditional formatting, hidden Rules sheet, and defined names.
