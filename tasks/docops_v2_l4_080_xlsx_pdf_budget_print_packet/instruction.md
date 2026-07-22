# Task

This task is part of the DocumentBenchmark L4 style/format expansion set.

## Goal

Create the required formatted outputs for the `budget print packet` package.

## Inputs

- `docops_v2_l4_080_xlsx_pdf_budget_print_packet_input.xlsx`: draft primary document with broken style/formatting
- `style_guide.docx`: authoritative colors, heading/table style guidance, and final status rules
- `source_rows.xlsx`: source rows to be carried into styled tables/registers where applicable
- `format_requirements.pdf`: authoritative cross-document formatting and cleanup requirements

## Required Outputs

Write all required files to `/root/submission`:

- `xlsxpdfsf_012_budget_print_packet_seed_styled_register.xlsx`
- `xlsxpdfsf_012_budget_print_packet_seed_styled_packet.pdf`

## Required Style/Format Work

- Remove all draft/scratch/manual-placeholder material.
- Apply primary color `7B2CBF` and accent color `F4A261` consistently.
- Preserve native format controls required by each output type: DOCX TOC/headings, PPTX native tables and reference slides, XLSX tables/print areas/validation/hidden rules, and PDF bookmarks/page order.
- Keep the outputs consistent with the style guide and source rows.
