            # Task

            This task is part of `harbor_tasks_l3_reddit_batch5_coupled`.

            ## Inputs

            - Source document: `/root/docops_v2_l3_010_docx_policy_exception_cascade_input.docx`
            - Original community-inspired seed description: `/root/original_task_description.txt`

            ## Goal

            A policy manual needs exceptions extracted into tables whose counts and blocking status affect the release header and control matrix.

            ## L3 Single-Document Workflow Instruction

            Remove the manual TOC placeholder; add a real TOC field; reorder sections to Scope, Controls, Exceptions, Appendix A - Definitions, Appendix B - Evidence; remove Draft wording; generate a Release Matrix linking CTRL-1 and CTRL-2 to their exception status; generate an Exception Register; highlight the MFA exception paragraph; start both appendices on new pages; set the header to include the open-exception count; normalize footer and styles.

            Keep the Release Matrix and Exception Register as native tables, not additional heading sections that interrupt the five required policy sections. The first native table should be the Release Matrix:

            - `Control | Requirement | Exception Owner | Release Status`
            - `CTRL-1 | MFA required | Unassigned | Blocked - MFA rollout missing owner`
            - `CTRL-2 | Vendor review required | Compliance team | Blocked - Vendor review overdue`

            The second native table should be the Exception Register:

            - `Exception | Owner | Due Date`
            - `MFA rollout missing owner | Unassigned | 2026-06-12`
            - `Vendor review overdue | Compliance team | 2026-06-14`

            Highlight the paragraph `Exception: MFA rollout missing owner.` The normalized header should be `Policy Manual | Exceptions Open: 2 | 2026-06-05`, and the footer should include `Page` without retaining confidential or draft wording.

            ## Output Requirements

            - Do not modify the source document in place.
            - Save the revised document to `/root/submission/docops_v2_l3_010_docx_policy_exception_cascade_output.docx`.
            - Keep the output format as `.docx`.
            - Complete the entire high-coupling workflow; partial local edits are not sufficient.

            ## Inspired By

            - Table of contents (https://www.reddit.com/r/MicrosoftWord/comments/1de3nu5)
- Create Table of contents for Annexes (https://www.reddit.com/r/MicrosoftWord/comments/1hsn1tb)
