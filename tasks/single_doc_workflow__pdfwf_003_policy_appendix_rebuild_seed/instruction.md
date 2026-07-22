# Task

This task is part of the DocumentBenchmark single-document workflow split.

## User Request

Rebuild a policy appendix PDF in one file.

## Primary Input

- `pdfwf_003_policy_appendix_rebuild_seed.pdf`

## Instructions

Remove visible Draft: prefixes, delete the obsolete appendix page, keep the Index page, restore the live pages to Cover, Index, Policy A, Policy B, and Appendix, and add bookmarks for each live page.

Please complete the request directly in the input document. Treat the file as a working draft with explicit draft markers. Remove the `Draft:` prefixes, remove obsolete scratch material, repair the requested structure, and otherwise keep the live wording unchanged.

Use the visible page titles to restore the requested page order; do not rename live pages.
Remove obsolete or superseded draft pages and remove explicit draft prefixes from the live pages.
Remove any superseded draft pages that are clearly not part of the live packet.

## Output

- Write the revised document to `/root/submission/pdfwf_003_policy_appendix_rebuild_seed.pdf`.
- Do not output a text-only answer; the required deliverable is the revised file.
