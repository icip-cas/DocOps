            # Task

            This task is part of `harbor_tasks_l3_reddit_batch4_over10`.

            ## Inputs

            - Source document: `/root/docops_v2_l3_030_pptx_webinar_over10_input.pptx`
            - Original community-inspired seed description: `/root/original_task_description.txt`

            ## Goal

            A webinar deck needs agenda, quiz, notes, numbering, and locked legal content synchronized after major cleanup.

            ## L3 Single-Document Workflow Instruction

            Delete Scratch Backup and Deprecated Poll; remove Draft from retained titles; reorder to Title, Agenda, Setup, Scenario, Deep Dive, Quiz, Actions, Reference - Legal; generate an Agenda native table with slide numbers; generate a Quiz native table; add section labels to delivery slides; add Slide X of 8 footer text to editable slides; add the required facilitator note to Actions; preserve Reference - Legal exactly; apply consistent release styling.

            The Agenda table must be a native PowerPoint table with exactly these rows:

            | Section | Slide |
            | --- | --- |
            | Setup | 3 |
            | Scenario | 4 |
            | Deep Dive | 5 |
            | Quiz | 6 |
            | Actions | 7 |

            The Quiz table must be a native PowerPoint table with exactly these rows:

            | Question | Required Action |
            | --- | --- |
            | How should participants respond to the live check-in? | Open poll |

            Add the exact section label `Section: Delivery` to Setup, Scenario, Deep Dive, Quiz, and Actions.

            Add these exact footer texts to the editable slides:

            - Title: `Slide 1 of 8`
            - Agenda: `Slide 2 of 8`
            - Setup: `Slide 3 of 8`
            - Scenario: `Slide 4 of 8`
            - Deep Dive: `Slide 5 of 8`
            - Quiz: `Slide 6 of 8`
            - Actions: `Slide 7 of 8`

            The Actions speaker notes must contain exactly:

            `Facilitator note: send recap within 24 hours.`

            ## Output Requirements

            - Do not modify the source document in place.
            - Save the revised document to `/root/submission/docops_v2_l3_030_pptx_webinar_over10_output.pptx`.
            - Keep the output format as `.pptx`.
            - Complete the entire high-coupling workflow; partial local edits are not sufficient.

            ## Inspired By

            - Slides not in Slide Master (https://www.reddit.com/r/powerpoint/comments/1rxet0f/slides_not_in_slide_master/)
- Custom slide numbering in PowerPoint (https://www.reddit.com/r/powerpoint/comments/w337xw)
