# Task

This task is part of the DocumentBenchmark L4 manual expansion set.

## Goal

Prepare the vendor onboarding tracker for release by completing required rows from multiple supporting documents.

## Primary Input

- `docops_v2_l4_044_excel_vendor_onboarding_tracker_from_policy_pack_input.xlsx`

## Supporting Inputs

- `vendor_evidence_policy_packet.pdf`
- `vendor_owner_directory.docx`
- `evidence_index.xlsx`

## Required Edit

Complete `Vendor Tracker` using the supporting policy packet, owner directory, and evidence index. Fill only the blank required vendor rows. Keep the optional `Delta Foods` row unchanged.

Use canonical owner names from `vendor_owner_directory.docx`, evidence IDs and source pages from the policy packet and evidence index, and exact review-note wording from the policy packet.

Set `Cover!B4` to `Ready for Release`.

## Output

- Write the revised workbook to `/root/submission/docops_v2_l4_044_excel_vendor_onboarding_tracker_from_policy_pack_output.xlsx`
- Preserve non-target structure and formatting unless explicitly asked to change it.

## Notes

- Do not hardcode release flags. Use live formulas in `Vendor Tracker!J2:J4` and `J6:J7`.
- Preserve `Summary!B3:B8` as formulas.
- Preserve the hidden `Rules` sheet and the `status_choices` defined name.
- Preserve the `Tracker Template` sheet unchanged.
- Preserve the native table range `Vendor Tracker!A1:J7`.
- Preserve data validation on `Vendor Tracker!F2:F7`.
- Highlight the rows whose final status is `Partial` or `Missing`.
