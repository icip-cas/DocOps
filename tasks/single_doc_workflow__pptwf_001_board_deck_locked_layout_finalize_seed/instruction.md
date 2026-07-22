# Task

This task is part of the DocumentBenchmark single-document workflow split.

## User Request

Finalize a board deck with a locked reference layout.

## Primary Input

- `pptwf_001_board_deck_locked_layout_finalize_seed.pptx`

## Instructions

Remove visible Draft: prefixes, delete the Draft Parking Lot slide, restore the live slide order from Board Memo through the locked reference slide, keep the Action Tracker table values unchanged, and do not modify the locked reference slide.

Please complete the request directly in the input document. Treat the file as a working draft with explicit draft markers. Remove the `Draft:` prefixes, remove obsolete scratch material, repair the requested structure, and otherwise keep the live wording unchanged.

Use the visible slide titles to restore the requested slide order; do not rename live slides.
Remove explicit draft prefixes from titles and bullets without paraphrasing the slide text.
Repair any slide table so it matches the surrounding action/checklist content and remains readable.
A slide labeled as a layout reference is locked; use it as a visual reference but do not modify that slide.
Keep the locked reference slide unchanged.
Remove any superseded draft slides before saving the final deck.

## Output

- Write the revised document to `/root/submission/pptwf_001_board_deck_locked_layout_finalize_seed.pptx`.
- Do not output a text-only answer; the required deliverable is the revised file.
