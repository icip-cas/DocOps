# Task

This task is part of a regenerated L3 single-document workflow set.

## Inputs

- Source PDF: `/root/docops_v2_l3_032_pdf_hazmat_field_guide_rebuild_input.pdf`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

Prepare a clean public hazardous-materials field guide from a messy internal working PDF.

## Single-PDF Workflow Instruction

Revise the source PDF into a public incident field guide. Complete the full workflow in one output `.pdf`:

- Save the revised PDF to `/root/submission/docops_v2_l3_032_pdf_hazmat_field_guide_rebuild_output.pdf`.
- Remove all visible `DRAFT`, `PRIVATE`, `INTERNAL`, `scratch`, and unverified material.
- Remove internal-only pages and rebuild the PDF with exactly these eight public pages in this order:
  1. `Capstone Harbor HazMat Field Guide`
  2. `Quick Actions`
  3. `Substance Index`
  4. `Protective Action Distances`
  5. `Isolation Zones`
  6. `Responder PPE Matrix`
  7. `Decontamination Workflow`
  8. `Public Release Log`
- Apply the public field-guide style: dark blue cover/header, terracotta quick-action header, gold reference-section headers, green operations-section headers, page footer, and clear page numbering.
- Apply all correction-log fixes:
  `Ammonia Anhydrus` to `Ammonia, anhydrous`, guide `128` to `125` for ammonia, chlorine small-spill night action `0.8 mi` to `1.1 mi`, gasoline initial isolation `100 ft` to `150 ft`, and release date `2026-07-01` to `2026-07-15`.
- Rebuild the substance index, protective action distances, isolation zones, PPE matrix, decontamination workflow, and public release log as clean public sections.
- Exclude chlorine plume rumor, vendor liability notes, personal phone, do-not-release language, draft evacuation language, and unverified appendix text.

## Atomic Scope

Only the requested single-PDF workflow edits are in scope. Preserve public field-guide facts, correction facts, section order, visual header bands, footer text, and PDF metadata requirements.

## Output Requirements

- Do not modify the source PDF in place.
- Keep the output format the same as the input PDF (`.pdf`).
- Complete the full PDF workflow, not just local text redaction.

## Expected Output Type

A revised `.pdf` artifact that completes a complex L3 single-document public field-guide rebuild, correction, cleanup, and style migration workflow.
