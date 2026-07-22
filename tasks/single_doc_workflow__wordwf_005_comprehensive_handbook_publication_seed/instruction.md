# Task

This task is part of the DocumentBenchmark single-document workflow split.

## User Request

Clean the comprehensive handbook draft into a structured DOCX without rewriting its live content.

## Primary Input

- `wordwf_005_comprehensive_handbook_publication_seed.docx`

## Instructions

Remove the visible Draft: prefixes, restore the exact chapter headings in order, refresh the table of contents, preserve the English and Chinese paragraphs, keep the table values unchanged, normalize body style, and put Appendix A - Release Checklist on a fresh page with the document header and footer.

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

- Write the revised document to `/root/submission/wordwf_005_comprehensive_handbook_publication_seed.docx`.
- Do not output a text-only answer; the required deliverable is the revised file.
