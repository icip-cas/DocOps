# Task

This task is part of a regenerated L3 single-document workflow set.

## Inputs

- Source presentation: `/root/docops_v2_l3_025_pptx_ai_incident_disclosure_packet_input.pptx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

Prepare a clean public AI incident disclosure briefing deck from a messy internal model-review presentation.

## Required Sequential Workflow

This task must be completed in order. Later steps depend on corrected facts from earlier steps:

1. Remove all internal-only slides and all visible or speaker-note confidential material.
2. Apply the correction log facts before rebuilding any public content.
3. Rebuild the public incident timeline from the corrected incident facts.
4. Rebuild the affected-services matrix from the corrected system names and impact counts.
5. Rebuild the remediation controls and transparency register from the corrected owners and dates.
6. Rebuild the disclosure checklist and board actions from the corrected timeline and control facts.
7. Apply the required public style migration after the public content has been rebuilt.
8. Preserve only concise public speaker notes.

## Single-Presentation Workflow Instruction

Revise the source presentation into a polished public AI incident disclosure packet. Complete the full workflow in one output `.pptx`:

- Save the revised presentation to `/root/submission/docops_v2_l3_025_pptx_ai_incident_disclosure_packet_output.pptx`.
- Remove all visible and speaker-note `DRAFT`, `PRIVATE`, `INTERNAL`, `scratch`, and confidential model-review material.
- Remove internal-only slides and rebuild the deck with exactly these nine public slides in this order:
  1. `CivicHelp AI Incident Disclosure Packet`
  2. `Disclosure Scope`
  3. `Corrected Incident Timeline`
  4. `Affected Services Matrix`
  5. `Remediation Controls`
  6. `Transparency Register`
  7. `Public Disclosure Checklist`
  8. `Board Action Items`
  9. `Publication Style Guide`
- Apply all correction-log fixes before rebuilding the public slides:
  `CivicHelp Chatbot v2.1-beta` to `CivicHelp Chatbot v2.2`, affected sessions `2,900` to `1,240`, first public notice `2026-04-03` to `2026-04-02`, owner `Jules Ren` to `Jules Renner`, and status `Containted` to `Contained`.

## Required Style Migration

The output must apply these explicit public-disclosure style requirements:

- Use title slide background color `111827`.
- Use content slide background color `F9FAFB`.
- Use Aptos Display bold titles on every slide.
- Use white title text on the title slide and `111827` title text on all content slides.
- Add an accent rectangle with fill color `2563EB` on every slide.
- Use public table header fill `374151` with white bold table-header text on every table slide.
- Add footer text `CivicHelp Public AI Disclosure | April 2026` on every slide.
- Keep concise public speaker notes on every slide, and remove confidential or personally identifying notes.

## Content Requirements

- Rebuild the timeline, affected-services matrix, remediation controls, transparency register, public disclosure checklist, board action items, and publication style guide as public slides.
- Exclude prompt transcript, jailbreak reproduction, vendor security exception, employee discipline note, personal phone, legal hold, do-not-publish language, and private model score.

## Output Requirements

- Do not modify the source presentation in place.
- Keep the output format the same as the input presentation (`.pptx`).
- Complete the full ordered workflow, not just local text replacement.

## Expected Output Type

A revised `.pptx` artifact that completes a strict-order L3 single-presentation AI incident disclosure workflow with correction, rebuilding, speaker-note cleanup, and explicit style migration.
