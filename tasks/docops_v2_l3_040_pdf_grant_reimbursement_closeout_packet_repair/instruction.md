# Grant Reimbursement Closeout Packet Repair

You are given a single PDF:

`docops_v2_l3_040_pdf_grant_reimbursement_closeout_packet_repair_input.pdf`

The source PDF is a draft grant reimbursement closeout packet with old award facts, unreconciled costs, payment-account details, identity fragments, staff-only hold notes, and missing release controls. Repair it into a public-safe closeout packet.

## Required Workflow

1. Remove old award facts, payment-account details, identity fragments, wire instructions, and draft/internal material.
2. Apply corrected facts everywhere:
   - Grant ID: `GRC-2026-118`
   - Recipient: `North Valley Food Bank`
   - Reporting period: `2026-04-01 to 2026-06-30`
   - Claim amount: `$184,620`
   - Eligible cost total: `$179,420`
   - Disallowed cost: `$5,200`
   - Final payable: `$179,420`
   - Public contact: `grants-closeout@state.example`
3. Rebuild exactly 6 Letter-size portrait pages in this order:
   - `Grant Reimbursement Closeout Packet`
   - `Corrected Award and Claim Summary`
   - `Eligible Cost Reconciliation`
   - `Evidence Crosswalk`
   - `Payment Hold Resolution Workflow`
   - `Final Certification Checklist`
4. Include the award summary, eligible cost reconciliation, evidence crosswalk, payment hold workflow, and certification checklist tables.
5. Remove all incorrect or restricted source remnants:
   - `DRAFT`
   - `INTERNAL ONLY`
   - `PRIVATE bank`
   - `GRC-2025-044`
   - `Lakeside Shelter`
   - `$192,800`
   - `routing number`
   - `021000021`
   - `wire instructions`
   - `SSN`
   - `account ending`
   - `grants-old@state.example`
6. Apply the final PDF style and controls:
   - Header bar color: `14532D`.
   - Accent color: `F59E0B`.
   - Footer on every page: `Grant Reimbursement Closeout Packet | GRC-2026-118`.
   - Bookmarks matching the final 6 page titles.
   - Metadata title: `Grant Reimbursement Closeout Packet GRC-2026-118`.
   - Metadata author: `State Grant Closeout Office`.
   - Metadata subject: `Corrected reimbursement closeout packet for GRC-2026-118`.

Save the completed PDF as:

`/root/submission/docops_v2_l3_040_pdf_grant_reimbursement_closeout_packet_repair_output.pdf`
