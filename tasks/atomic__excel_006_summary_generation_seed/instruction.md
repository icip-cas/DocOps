# Task

This task is adapted from the DocumentBenchmark seed `excel_006`.

## Inputs

- Source document: `/root/excel_006_summary_generation_seed.xlsx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This workbook has a raw incident log and a request for a concise executive summary sheet.

## Single-Step Benchmark Instruction

Create a new sheet named `Executive Summary` and generate three short summary bullets that synthesize the main operational themes from the `Incident Log` sheet.

## Atomic Scope

Only content generation is explicitly requested in this task.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/excel_006_summary_generation_seed.xlsx`.
- Keep the output format the same as the input document (`.xlsx`).
- Focus on the requested atomic operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .xlsx workbook containing the new generated summary sheet.

## Why This Variant Is Hard

- The correct output is not a direct copy of any one row
- Multiple incidents overlap across categories and sites
- Good answers require abstraction, not just counting

## Inspired By

- Need a way to get avg times for different factors. (https://www.reddit.com/r/excel/comments/1sjx8f9/need_a_way_to_get_avg_times_for_different_factors/)
