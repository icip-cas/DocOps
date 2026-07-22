# Task

This task is part of the DocumentBenchmark single-document workflow split.

## User Request

Clean the board packet PDF while preserving the live page text.

## Primary Input

- `pdfwf_004_board_packet_cover_refresh_seed.pdf`

## Instructions

Remove visible Draft: prefixes, delete the draft cover page, restore the live pages to Cover, Management Summary, Working Pages, and Closing, and add bookmarks for each live page.

Please complete the request directly in the input document. Treat the file as a working draft with explicit draft markers. Remove the `Draft:` prefixes, remove obsolete scratch material, repair the requested structure, and otherwise keep the live wording unchanged.

Use the visible page titles to restore the requested page order; do not rename live pages.
Remove obsolete or superseded draft pages and remove explicit draft prefixes from the live pages.
Remove any superseded draft pages that are clearly not part of the live packet.

## Output

- Write the revised document to `/root/submission/pdfwf_004_board_packet_cover_refresh_seed.pdf`.
- Do not output a text-only answer; the required deliverable is the revised file.
