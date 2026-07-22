# Task

This task is part of the DocumentBenchmark single-document workflow split.

## User Request

Rebuild a closeout binder PDF.

## Primary Input

- `pdfwf_005_closeout_binder_page_labels_and_bookmarks_seed.pdf`

## Instructions

Remove visible Draft: prefixes, delete the obsolete divider page, restore the live pages to Cover, Meeting Log, Binder Index, Approved Exhibit, Appendix A, and Appendix B, and add bookmarks for each live page.

Please complete the request directly in the input document. Treat the file as a working draft with explicit draft markers. Remove the `Draft:` prefixes, remove obsolete scratch material, repair the requested structure, and otherwise keep the live wording unchanged.

Use the visible page titles to restore the requested page order; do not rename live pages.
Remove obsolete or superseded draft pages and remove explicit draft prefixes from the live pages.
Remove any superseded draft pages that are clearly not part of the live packet.

## Output

- Write the revised document to `/root/submission/pdfwf_005_closeout_binder_page_labels_and_bookmarks_seed.pdf`.
- Do not output a text-only answer; the required deliverable is the revised file.
