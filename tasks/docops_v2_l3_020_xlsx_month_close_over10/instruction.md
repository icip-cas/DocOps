            # Task

            This task is part of `harbor_tasks_l3_reddit_batch4_over10`.

            ## Inputs

            - Source document: `/root/docops_v2_l3_020_xlsx_month_close_over10_input.xlsx`
            - Original community-inspired seed description: `/root/original_task_description.txt`

            ## Goal

            A monthly close workbook needs a release-ready summary driven by cleaned transactions and exception logic.

            ## L3 Single-Document Workflow Instruction

            Create Clean Transactions and Exceptions sheets; clean and reorder the raw transactions; add Category, Owner, Receipt Required, Review Flag, and Signed Amount formulas; convert cleaned data to CleanTransactions; add Receipt validation; highlight review rows; generate exception rows; repair Close Summary revenue, expense, and net formulas; rename the summary title; hide Raw Transactions and Archive; reorder sheets for release; apply basic release formatting.

            Use the release sheet order `Close Summary`, `Clean Transactions`, `Exceptions`,
            `Raw Transactions`, `Vendor Map`, `Rules`, `Archive`. `Clean Transactions` must keep
            all seven raw transaction rows after cleaning/reordering and use the columns `Date`,
            `Vendor`, `Amount`, `Receipt`, `Memo`, `Category`, `Owner`, `Receipt Required`,
            `Review Flag`, and `Signed Amount`. Keep `Category`, `Owner`, `Receipt Required`,
            `Review Flag`, `Signed Amount`, and the Close Summary revenue/expense/net metrics
            formula-backed. `Exceptions` should include one row per review transaction, including
            missing receipts, duplicate transactions, unmapped vendors, and threshold exceptions.

            ## Output Requirements

            - Do not modify the source document in place.
            - Save the revised document to `/root/submission/docops_v2_l3_020_xlsx_month_close_over10_output.xlsx`.
            - Keep the output format as `.xlsx`.
            - Complete the entire high-coupling workflow; partial local edits are not sufficient.

            ## Inspired By

            - Trying to find unique values when comparing hundreds of millions of rows (https://np.reddit.com/r/excel/comments/uoixte/trying_to_find_unique_values_when_comparing/)
- Index Match to return multiple values (https://us.reddit.com/r/excel/comments/seajec/how_can_i_use_index_match_to_return_multiple/)
