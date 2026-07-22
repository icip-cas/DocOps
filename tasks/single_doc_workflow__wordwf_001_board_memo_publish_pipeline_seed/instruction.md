# Task

This task is part of the DocumentBenchmark single-document workflow split.

## User Request

Clean the board memorandum draft without changing the live wording.

## Primary Input

- `wordwf_001_board_memo_publish_pipeline_seed.docx`

## Instructions

Remove the visible Draft: prefixes, restore the sections to Executive Summary, Program Snapshot, Board Actions, and Appendix A - Metrics, refresh the table of contents, keep the table values unchanged, and normalize the header and footer.

Please complete the request directly in the input document. Treat the file as a working draft with explicit draft markers. Remove the `Draft:` prefixes, remove obsolete scratch material, repair the requested structure, and otherwise keep the live wording unchanged.

Use the visible section titles to restore the document order; do not rename sections or paraphrase body paragraphs.
Clean only explicit draft markers and obsolete scratch notes.
Repair the table by removing draft prefixes from cells and preserving the visible cell values.
Replace the placeholder table of contents with a real Word table-of-contents field.
Move appendix-style material onto a fresh page or section where appropriate.
Normalize the header and footer so the document looks ready for circulation.
Do not add numeric prefixes to section headings unless the source already uses that exact heading text.
Do not leave provisional reviewer notes, draft-only sections, or placeholder navigation text in the final file.

## Output

- Write the revised document to `/root/submission/wordwf_001_board_memo_publish_pipeline_seed.docx`.
- Do not output a text-only answer; the required deliverable is the revised file.
