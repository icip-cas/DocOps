            # Task

            This task is part of `harbor_tasks_l3_longflow_batch1`.

            ## Inputs

            - Source document: `/root/docops_v2_l3_011_xlsx_service_recovery_closeout_input.xlsx`
            - Original community-inspired seed description: `/root/original_task_description.txt`

            ## Goal

            A field-service closeout workbook needs a long, ordered release workflow: raw operational tickets must become a controlled service recovery pack with formulas, exceptions, customer notices, dashboard metrics, and release-safe sheet structure.

            ## L3 Single-Document Long-Flow Instruction

            Starting from the source workbook, perform the complete workflow in order:

            1. Replace the stale `Legacy Dashboard` with a new `Dashboard`.
            2. Create `Release Queue`, `Exception Review`, and `Customer Notices`.
            3. Keep `Raw Intake`, `SLA Policy`, `Owner Map`, `Part Catalog`, and `Technician Roster` values unchanged.
            4. Populate `Release Queue` from all nine rows in `Raw Intake`.
            5. Add the columns `Priority`, `Owner`, `SLA Hours`, `Age Hours`, `SLA Breach`, `Part Cost`, `Escalation Reason`, `Publish Note`, `Customer Notice`, and `Ready To Publish`.
            6. Infer priority from tier and issue type: Platinum or Power cases are P1, Gold is P2, Silver is P3, and Bronze is P4.
            7. Map owners by region from `Owner Map`.
            8. Map SLA hours by priority from `SLA Policy`.
            9. Use `Dashboard!B2` as the fixed review timestamp for open items when computing age hours.
            10. Flag SLA breaches when age hours exceed SLA hours, including closed items that closed late.
            11. Map part costs from `Part Catalog`.
            12. Generate escalation reasons from breach/parts/open/closed status.
            13. Suppress any internal "Do not publish" notes before customer-facing output.
            14. Generate customer notice text from customer, reason, and owner.
            15. Create `Exception Review` for breached, waiting-on-parts, or sensitive-note rows.
            16. Create `Customer Notices` only for rows ready to publish; do not include the waiting-on-parts sensitive row.
            17. Build dashboard formulas for open tickets, waiting-on-parts tickets, SLA breaches, ready notices, total part exposure, and West-region breaches.
            18. Convert the release, exception, and notice ranges into native Excel tables.
            19. Add status data validation to `Release Queue!G2:G10`, freeze the release queue header, and highlight flagged release rows.
            20. Reorder sheets for release and hide `Raw Intake`, `Scratch Notes`, and `Archive`.

            ## Output Requirements

            - Do not modify the source document in place.
            - Save the revised document to `/root/submission/docops_v2_l3_011_xlsx_service_recovery_closeout_output.xlsx`.
            - Keep the output format as `.xlsx`.
            - Complete the full long-flow sequence; a partial cleanup or dashboard-only update is insufficient.

            ## Inspired By

            - Real-world Excel cleanup workflows involving lookup tables, data validation, conditional formatting, and dashboard repair.
            - Community spreadsheet pain around multi-sheet operational reports and publishable workbook release packs.
