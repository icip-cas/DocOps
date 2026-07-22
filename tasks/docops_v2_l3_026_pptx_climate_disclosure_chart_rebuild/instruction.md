# Task

This task is part of a regenerated L3 single-document workflow set.

## Inputs

- Source presentation: `/root/docops_v2_l3_026_pptx_climate_disclosure_chart_rebuild_input.pptx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

Prepare a clean public climate disclosure briefing deck from a messy internal sustainability review presentation.

## Required Sequential Workflow

This task must be completed in order:

1. Remove internal-only slides and confidential speaker-note material.
2. Apply the correction-log facts before rebuilding any public table or chart.
3. Rebuild the emissions inventory table using the corrected Scope 2 and renewable-energy values.
4. Rebuild the emissions trend chart from the corrected annual emissions values.
5. Rebuild the assurance status, transition actions, and board actions from the corrected facts.
6. Apply the required public visual style after public content and chart data are correct.
7. Preserve only concise public speaker notes.

## Single-Presentation Workflow Instruction

Revise the source presentation into a polished public climate disclosure briefing deck. Complete the full workflow in one output `.pptx`:

- Save the revised presentation to `/root/submission/docops_v2_l3_026_pptx_climate_disclosure_chart_rebuild_output.pptx`.
- Remove all visible and speaker-note `DRAFT`, `PRIVATE`, `INTERNAL`, `scratch`, and confidential sustainability-review material.
- Remove internal-only slides and rebuild the deck with exactly these nine public slides in this order:
  1. `Northline Climate Disclosure Briefing`
  2. `Disclosure Boundary`
  3. `Corrected Emissions Inventory`
  4. `Emissions Trend Chart`
  5. `Assurance Status`
  6. `Transition Action Plan`
  7. `Supplier Engagement Register`
  8. `Board Actions`
  9. `Publication Style Guide`
- Apply all correction-log fixes before rebuilding public slides:
  `Scope 2 location-based 14,800` to `13,480`, renewable electricity `42%` to `48%`, base year `2023` to `2024`, assurance owner `Kara Sato` to `Kara Saito`, and status `Submittd` to `Submitted`.

## Required Chart Feature

- Rebuild a clustered column chart on `Emissions Trend Chart`.
- The chart must have categories `2024`, `2025`, `2026`.
- The chart must have one series named `Total emissions tCO2e`.
- The chart values must be `42800`, `39750`, and `36220`.

## Required Style Migration

- Use title slide background color `0F172A`.
- Use content slide background color `F8FAFC`.
- Use Aptos Display bold titles on every slide.
- Use white title text on the title slide and `0F172A` title text on all content slides.
- Add an accent rectangle with fill color `22C55E` on every slide.
- Use public table header fill `166534` with white bold table-header text on every table slide.
- Add footer text `Northline Public Climate Disclosure | FY2026` on every slide.
- Keep concise public speaker notes on every slide, and remove confidential or personally identifying notes.

## Content Requirements

- Rebuild the boundary, emissions inventory, assurance status, transition action plan, supplier engagement register, board actions, and publication style guide as public slides.
- Exclude unreleased forecast, facility closure rumor, supplier dispute, personal phone, legal hold, do-not-publish language, and private assurance finding.

## Output Requirements

- Do not modify the source presentation in place.
- Keep the output format the same as the input presentation (`.pptx`).
- Complete the full ordered workflow, not just local text replacement.

## Expected Output Type

A revised `.pptx` artifact that completes a strict-order L3 single-presentation climate disclosure workflow with chart rebuilding, table rebuilding, speaker-note cleanup, and explicit style migration.
