            # Task

            This task is part of `harbor_tasks_l3_longflow_batch1`.

            ## Inputs

            - Source document: `/root/docops_v2_l3_029_pptx_qbr_release_deck_input.pptx`
            - Original community-inspired seed description: `/root/original_task_description.txt`

            ## Goal

            Turn a draft QBR deck into a release-ready presentation through a long ordered deck publication workflow.

            ## L3 Single-Document Long-Flow Instruction

            Complete the full deck release workflow:

            1. Delete `Scratch Backup` and `Deprecated Roadmap`.
            2. Remove all `Draft:` prefixes and working-copy language.
            3. Reorder the deck to: Title, Agenda, Market Context, KPI Snapshot, Customer Risks, Mitigation Plan, Decision Request, Appendix - Data Notes, Reference - Locked Legal.
            4. Rename the opening slide to `Title`.
            5. Rebuild the agenda table with Section, Slide, and Owner columns.
            6. Synchronize agenda slide numbers to the final 9-slide deck.
            7. Add `QBR Release | 2026-06-05` and `Slide X of 9` to every content slide except Title, Agenda, and the locked legal reference slide.
            8. Update KPI Snapshot to ARR 3.2M, churn risk 4 accounts, NRR 94%, and expansion pipeline 1.1M.
            9. Add a QBR readout column to the KPI table.
            10. Replace the customer risk table with renewal delay, adoption gap, and support escalation rows.
            11. Create a mitigation plan table with action, owner, due date, and status.
            12. Rewrite Decision Request with the required three decision statements.
            13. Add the required Decision Request speaker note.
            14. Rewrite Appendix - Data Notes with the public source sentence.
            15. Remove internal escalation text from publishable slides.
            16. Preserve `Reference - Locked Legal` text, notes, and shape count exactly as in the input.
            17. Save as `.pptx`.

            Use the following release content specifications for the rebuilt tables and rewritten slides.

            Agenda table:

            | Section | Slide | Owner |
            | --- | --- | --- |
            | Market Context | 3 | Maya Chen |
            | KPI Snapshot | 4 | Jon Bell |
            | Customer Risks | 5 | Priya Rao |
            | Mitigation Plan | 6 | Maya Chen |
            | Decision Request | 7 | Jon Bell |
            | Appendix - Data Notes | 8 | Priya Rao |

            KPI Snapshot table:

            | Metric | Value | QBR Readout |
            | --- | --- | --- |
            | ARR | 3.2M | Up 10% QoQ |
            | Churn risk | 4 accounts | Requires executive review |
            | NRR | 94% | Below target |
            | Expansion pipeline | 1.1M | Two deals require support plan |

            Customer Risks table:

            | Risk | Account Count | Owner | Mitigation |
            | --- | --- | --- | --- |
            | Renewal delay | 2 | Maya Chen | Executive outreach by 2026-06-10 |
            | Adoption gap | 1 | Priya Rao | Enablement plan by 2026-06-12 |
            | Support escalation | 1 | Jon Bell | Daily incident review |

            Mitigation Plan table:

            | Action | Owner | Due Date | Status |
            | --- | --- | --- | --- |
            | Confirm renewal-risk owners | Maya Chen | 2026-06-07 | Open |
            | Publish enablement plan | Priya Rao | 2026-06-12 | Open |
            | Start daily incident review | Jon Bell | 2026-06-06 | Ready |

            Decision Request must contain these three statements:

            - `Approve executive outreach for 4 renewal-risk accounts.`
            - `Authorize enablement plan for NRR recovery.`
            - `Decision needed by 2026-06-07.`

            Decision Request speaker notes must contain exactly:

            `Speaker note: ask for decision on renewal-risk outreach before reviewing appendix.`

            Appendix - Data Notes must contain:

            `Public source: QBR metrics workbook, 2026-06-05.`

            ## Output Requirements

            - Do not modify the source document in place.
            - Save the revised deck to `/root/submission/docops_v2_l3_029_pptx_qbr_release_deck_output.pptx`.
            - Keep the output format as `.pptx`.
            - Complete the full long-flow release workflow; isolated slide edits are insufficient.

            ## Inspired By

            - Real QBR deck release pain involving stale agenda tables, draft slides, page numbers, metrics tables, action plans, speaker notes, and locked reference slides.
