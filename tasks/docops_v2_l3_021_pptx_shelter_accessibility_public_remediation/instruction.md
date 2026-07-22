# Task

This task is part of a regenerated L3 single-document workflow set.

## Inputs

- Source presentation: `/root/docops_v2_l3_021_pptx_shelter_accessibility_public_remediation_input.pptx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

Prepare a clean accessible public emergency shelter briefing deck from a messy internal accessibility review presentation.

## Required Sequential Workflow

This task must be completed in order:

1. Remove internal-only slides, hidden off-slide content, confidential notes, and staff-only links.
2. Apply the correction memo before rebuilding public counts, tables, contact details, and update log.
3. Rebuild the public shelter accessibility briefing in the required slide order.
4. Add meaningful alt text to the key visual on every slide.
5. Replace broken or internal hyperlinks with public links.
6. Apply the required visual style, footer, speaker-note cleanup, and table styling.

## Single-Presentation Workflow Instruction

Revise the source presentation into one clean public `.pptx`:

- Save the revised presentation to `/root/submission/docops_v2_l3_021_pptx_shelter_accessibility_public_remediation_output.pptx`.
- Remove visible and hidden `DRAFT`, `PRIVATE`, `INTERNAL`, staff-only, off-slide, personal phone, emergency operations center, and do-not-publish material.
- Remove internal-only slides and rebuild the deck with exactly these eight public slides in this order:
  1. `Emergency Shelter Public Briefing`
  2. `Corrected Shelter Network`
  3. `Accessibility Map`
  4. `Transit Access Updates`
  5. `Service Animal and Medical Equipment`
  6. `Hotline and Web Links`
  7. `Update Log`
  8. `Accessibility Checklist`

## Required Corrections

Apply these corrections before rebuilding public slides:

- Shelter count must be `18`, not `16`.
- Wheelchair accessible shelter count must be `14`, not `11`.
- Public hotline must be `1-800-555-0148`, not a staff personal phone.
- Public website must be `https://city.example/shelters/public`, not an internal URL.
- Public update date must be `2026-07-20`.

## Required Accessibility and Link Features

- Add meaningful alt text to the key visual on every slide.
- On `Hotline and Web Links`, include a hyperlink to `https://city.example/shelters/public`.
- On `Hotline and Web Links`, include a `mailto:` hyperlink to `accessibility@city.example`.
- Do not leave broken, internal, or staff-only hyperlinks.

## Required Style Migration

- Use title slide background color `102A43`.
- Use content slide background color `F7FBFF`.
- Use `Aptos Display` bold titles on every slide.
- Use white title text on the title slide and `102A43` title text on all content slides.
- Add an accent rectangle with fill color `F59E0B` on every slide.
- Use public table header fill `0B5CAD` with white bold table-header text.
- Add footer text `Public Shelter Accessibility Briefing | 2026` on every slide.
- Keep concise public speaker notes on every slide, and remove confidential or staff-only notes.

## Output Requirements

- Do not modify the source presentation in place.
- Keep the output format the same as the input presentation (`.pptx`).
- Complete the full ordered workflow, not just local text replacement.

## Expected Output Type

A revised `.pptx` artifact that completes a strict-order L3 single-presentation accessibility remediation workflow with correction application, hidden-content cleanup, alt-text remediation, hyperlink repair, speaker-note cleanup, and style migration.
