# Task

This task is adapted from the DocumentBenchmark seed `pdf_002`.

## Inputs

- Source document: `/root/pdf_002_contract_conflict_reasoning_seed.pdf`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This contract excerpt contains payment clauses and schedule notes that may not agree with each other.

## Single-Step Benchmark Instruction

Identify the conflict in the payment terms and explain which clauses or schedule entries are inconsistent with one another.

## Atomic Scope

Only reasoning over internal document consistency is explicitly requested in this task.

## Output Requirements

- Do not modify the source document in place.
- Write your final answer to `/root/submission/final_answer.txt`.
- Keep the answer concise, directly responsive, and grounded in the provided document.

## Expected Output Type

A concise explanation of the conflicting provisions.

## Latent Issues Intentionally Left In The Seed

- The PDF mixes clause text with schedule summaries
- There is no explicit correction note inside the contract body
- The layout is slightly messy and could be reorganized, but that is not the explicit request

## Inspired By

- How do you edit a pdf document while keeping everything the same font and layout exactly? (https://www.reddit.com/r/pdf/comments/1snrzd5/how_do_you_edit_a_pdf_document_while_keeping/)
- Automatic Redaction Tools (https://www.reddit.com/r/pdf/comments/1s3kg1q/automatic_redaction_tools/)
