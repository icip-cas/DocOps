            # Task

            This task is part of `harbor_tasks_l3_longflow_batch1`.

            ## Inputs

            - Source document: `/root/docops_v2_l3_009_docx_incident_manual_publication_input.docx`
            - Original community-inspired seed description: `/root/original_task_description.txt`

            ## Goal

            Convert a draft incident response manual into a release-safe public Word manual through a long ordered workflow.

            ## L3 Single-Document Long-Flow Instruction

            Complete the full publication workflow:

            1. Remove the draft title and replace it with a public release title.
            2. Replace the manual TOC placeholder with a real Word TOC field.
            3. Remove all draft prefixes from section headings.
            4. Reorder sections into: Publication Summary, Response Timeline, Customer Communication, Control Exceptions, Action Register, Appendix A - Evidence Index, Appendix B - Approval Record.
            5. Remove internal privileged root-cause notes and phone bridge details.
            6. Normalize the response timeline into a public 3-column table.
            7. Create a customer communication table with message owners and ready statuses.
            8. Rewrite control exceptions into a public exception table.
            9. Create an action register table from the unresolved exception.
            10. Move evidence before approvals.
            11. Convert evidence into public labels and disposition.
            12. Convert approvals into a final approval record.
            13. Insert page breaks before Control Exceptions, Action Register, Appendix A, and Appendix B.
            14. Add a public release header.
            15. Add a page-number footer.
            16. Highlight the open MFA evidence exception paragraph.
            17. Ensure no "Draft:", "[INTERNAL]", "Do not publish", phone bridge, or privileged vendor-root-cause text remains.
            18. Preserve the document as `.docx`.

            ## Public Release Specification

            Use the following public release content for the transformed manual:

            - Title: `Incident Response Manual | Public Release`
            - Header: `Incident Response Manual | Public Release | 2026-06-05`
            - Footer: include page numbering with the label `Page`
            - Publication Summary must include:
              - `This public manual summarizes the June 2026 incident response without privileged root-cause analysis.`
              - `Customer-facing language has been normalized and internal notes have been removed.`
            - Response Timeline table:
              - `Time | Milestone | Public Summary`
              - `2026-06-01 08:12 | Alert opened | Incident response process started.`
              - `2026-06-01 09:40 | Containment applied | Containment controls were applied.`
              - `2026-06-01 15:10 | Service restored | Customer service was restored.`
            - Customer Communication table:
              - `Audience | Message Owner | Required Update | Status`
              - `Customers | Maya Chen | Send public restoration summary. | Ready`
              - `Support desk | Jon Bell | Use approved FAQ language only. | Ready`
            - Control Exceptions must include the paragraphs:
              - `Exception: MFA evidence owner remains unassigned.`
              - `Resolved exception: log retention evidence was exported and approved.`
            - Control Exceptions table:
              - `Exception ID | Control | Owner | Release Status`
              - `EX-01 | MFA evidence | Unassigned | Open - owner required`
              - `EX-02 | Log retention | Rina Patel | Closed - evidence exported`
            - Action Register table:
              - `Action ID | Owner | Due Date | Publish Status`
              - `ACT-01 | Unassigned | 2026-06-07 | Blocked until owner assigned`
              - `ACT-02 | Maya Chen | 2026-06-06 | Ready`
            - Evidence Index table:
              - `Evidence ID | Public Label | Disposition`
              - `EV-01 | Alert chronology | Publish`
              - `EV-02 | Internal Slack excerpt | Exclude`
              - `EV-03 | Customer notification record | Publish`
            - Approval Record table:
              - `Approver | Role | Decision`
              - `Maya Chen | Incident Lead | Approved`
              - `Jon Bell | Legal | Approved after privileged text removal`
            - Highlight the paragraph `Exception: MFA evidence owner remains unassigned.`

            ## Output Requirements

            - Do not modify the source document in place.
            - Save the revised document to `/root/submission/docops_v2_l3_009_docx_incident_manual_publication_output.docx`.
            - Keep the output format as `.docx`.
            - Complete the entire long-flow publication workflow; partial section cleanup is not sufficient.

            ## Inspired By

            - Long Word document publication pain involving TOCs, section breaks, headers/footers, appendices, and public/internal content boundaries.
