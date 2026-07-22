# Task

This task is adapted from the DocumentBenchmark seed `ppt_009`.

## Inputs

- Source document: `/root/ppt_009_metric_conflict_reasoning_seed.pptx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This deck contains KPI commentary that may not be internally consistent.

## Single-Step Benchmark Instruction

Identify the inconsistency in the delay reporting and explain why the slides do not fully agree.

## Atomic Scope

Only reasoning over internal consistency is explicitly requested in this task.

## Output Requirements

- Do not modify the source document in place.
- Write your final answer to `/root/submission/final_answer.txt`.
- Keep the answer concise, directly responsive, and grounded in the provided document.

## Expected Output Type

A concise explanation of the inconsistency.

## Why This Variant Is Hard

- The inconsistency depends on baseline interpretation, not just number mismatch
- All slides sound plausible on their own
- The correct explanation requires cross-slide reasoning
