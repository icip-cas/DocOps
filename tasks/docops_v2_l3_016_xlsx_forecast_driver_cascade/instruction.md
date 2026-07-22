            # Task

            This task is part of `harbor_tasks_l3_reddit_batch5_coupled`.

            ## Inputs

            - Source document: `/root/docops_v2_l3_016_xlsx_forecast_driver_cascade_input.xlsx`
            - Original community-inspired seed description: `/root/original_task_description.txt`

            ## Goal

            A forecast workbook needs a scenario driver that cascades through model formulas, exception review, and dashboard totals.

            ## L3 Single-Document Workflow Instruction

            Create Forecast Model and Exception Review; use Assumptions scenario validation to drive growth, price, revenue, capacity, and capacity flags; convert the model to a native table; highlight capacity breaches; update dashboard title and formulas from the model; generate exception rows from flagged months; hide Demand Raw and Archive; reorder sheets for release; preserve assumptions as the editable driver.

            Use the release sheet order `Dashboard`, `Forecast Model`, `Exception Review`,
            `Assumptions`, `Demand Raw`, `Archive`. `Assumptions!B5` should remain the editable
            scenario driver with validation for the available scenarios. `Forecast Model` must
            be a native table named `ForecastModel`, keep all four raw demand months, and include
            the columns `Month`, `Base Units`, `Actual Units`, `Scenario Growth`,
            `Forecast Units`, `Price`, `Revenue`, `Capacity`, and `Capacity Flag`.
            `Scenario Growth`, `Forecast Units`, `Price`, `Revenue`, `Capacity`, and
            `Capacity Flag` must remain formula-backed from the scenario driver and demand rows.
            `Exception Review` should include one row per capacity breach month, with fields that
            identify the month, the capacity issue, forecast units, capacity, and overage. The
            dashboard title should clearly be a forecast dashboard for the selected scenario, and
            dashboard totals/counts must remain formula-backed from `Forecast Model`.

            ## Output Requirements

            - Do not modify the source document in place.
            - Save the revised document to `/root/submission/docops_v2_l3_016_xlsx_forecast_driver_cascade_output.xlsx`.
            - Keep the output format as `.xlsx`.
            - Complete the entire high-coupling workflow; partial local edits are not sufficient.

            ## Inspired By

            - Excel formula refresh and linked outputs (https://dd.reddit.com/r/excel/comments/zwdlyx/can_i_add_a_currency_format_to_a_formula/)
- Index Match formulas in shared workbooks (https://us.reddit.com/r/excel/comments/seajec/how_can_i_use_index_match_to_return_multiple/)
