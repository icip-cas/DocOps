# Task

This task is part of a regenerated L3 single-document workflow set.

## Inputs

- Source presentation: `/root/docops_v2_l3_024_pptx_hospital_evacuation_public_briefing_input.pptx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

Prepare a clean public briefing deck from a messy internal hospital evacuation exercise presentation.

## Single-Presentation Workflow Instruction

Revise the source presentation into a polished public exercise briefing deck. Complete the full workflow in one output `.pptx`:

- Save the revised presentation to `/root/submission/docops_v2_l3_024_pptx_hospital_evacuation_public_briefing_output.pptx`.
- Remove all visible and speaker-note `DRAFT`, `PRIVATE`, `INTERNAL`, `scratch`, and confidential patient-safety review material.
- Remove internal-only slides and rebuild the deck with exactly these ten public slides in this order:
  1. `Harborview Hospital Evacuation Exercise: Public Briefing`
  2. `Exercise Scope and Assumptions`
  3. `Patient Movement Timeline`
  4. `Unit Impact Matrix`
  5. `Transport Resource Plan`
  6. `Family Reunification Workflow`
  7. `Communications Cadence`
  8. `Corrective Action Commitments`
  9. `Publication Style Guide`
  10. `Appendix Slide Index`
- Apply all correction-log fixes:
  `West Tower 5B` to `West Tower 5C`, ambulance count `14` to `18`, family center opening `10:45` to `10:15`, `Peds Unit` to `Pediatric Unit`, and `Dr. Helen Mora` to `Dr. Helena Mora`.

## Required Style Migration

The output must apply these explicit public-briefing style requirements:

- Use title slide background color `172A3A`.
- Use content slide background color `F8F4EC`.
- Use Georgia bold titles on every slide.
- Use white title text on the title slide and `172A3A` title text on all content slides.
- Add an accent rectangle with fill color `E4572E` on every slide.
- Use public table header fill `0B6E4F` with white bold table-header text on every table slide.
- Add footer text `Harborview Public Exercise Briefing | September 2026` on every slide.
- Keep concise public speaker notes on every slide, and remove confidential or personally identifying notes.

## Content Requirements

- Rebuild the timeline, unit impact matrix, transport resource plan, family reunification workflow, communications cadence, corrective action commitments, publication style guide, and appendix index as public slides.
- Exclude patient name, nurse grievance, insurer notification, legal risk memo, bed shortage rumor, security camera note, personal phone, and do-not-disclose language.

## Atomic Scope

Only the requested single-presentation workflow edits are in scope. Preserve public exercise facts, correction facts, slide order, speaker-note cleanup, table style, footer text, and visual style requirements.

## Output Requirements

- Do not modify the source presentation in place.
- Keep the output format the same as the input presentation (`.pptx`).
- Complete the full deck workflow, not just local text replacement.

## Expected Output Type

A revised `.pptx` artifact that completes a high-complexity L3 single-presentation public briefing rebuild, correction, cleanup, speaker-note, and explicit style migration workflow.
