# Task

This task is part of the DocumentBenchmark single-document workflow split.

## User Request

Clean up a team sync deck with caption fixes.

## Primary Input

- `pptwf_005_team_sync_deck_caption_cleanup_seed.pptx`

## Instructions

Remove visible Draft: prefixes, delete the Old Roadmap Draft slide, restore the live slide order from Team Sync through the locked reference slide, and do not modify the locked reference slide.

Please complete the request directly in the input document. Treat the file as a working draft with explicit draft markers. Remove the `Draft:` prefixes, remove obsolete scratch material, repair the requested structure, and otherwise keep the live wording unchanged.

Use the visible slide titles to restore the requested slide order; do not rename live slides.
Remove explicit draft prefixes from titles and bullets without paraphrasing the slide text.
A slide labeled as a layout reference is locked; use it as a visual reference but do not modify that slide.
Keep the locked reference slide unchanged.
Remove any superseded draft slides before saving the final deck.

## Output

- Write the revised document to `/root/submission/pptwf_005_team_sync_deck_caption_cleanup_seed.pptx`.
- Do not output a text-only answer; the required deliverable is the revised file.
