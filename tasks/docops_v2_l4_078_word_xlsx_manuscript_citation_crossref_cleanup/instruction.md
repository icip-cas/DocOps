# Task

Create the revised manuscript and citation/cross-reference audit workbook.

## Inputs

- `docops_v2_l4_078_word_xlsx_manuscript_citation_crossref_cleanup_input.docx`
- `draft_manuscript_author_date.docx`
- `journal_submission_guide.pdf`
- `reviewer_cleanup_requests.pdf`
- `reference_library.xlsx`
- `figure_table_inventory.docx`

## Required Outputs

Write both files to `/root/submission`:

- `revised_manuscript.docx`
- `citation_crossref_audit.xlsx`

## Required Work

- Convert author-date citations to numeric Vancouver-style citations in first-appearance order.
- Rebuild the reference list from the reference library in the new citation order.
- Fix figure and table references after removed draft items.
- Remove obsolete Supplementary Figure S1 main-text callout.
- Build a citation/cross-reference audit workbook with native Excel tables, formulas, data validation, print areas, defined names, and hidden private notes.
- Exclude internal author notes and reviewer-sensitive comments from public outputs.
