# Task

This task is adapted from the DocumentBenchmark seed `ppt_008`.

## Inputs

- Source document: `/root/ppt_008_bullet_editing_seed.pptx`
- Original benchmark seed description: `/root/original_task_description.txt`

## Goal

This client-facing slide contains bullets that are too blunt and emotionally charged.

## Single-Step Benchmark Instruction

Rewrite the three draft bullets so they sound professional and executive-appropriate while preserving the underlying message slots:
- paperwork/files still need correction
- corrected files/materials are needed today or promptly
- another delay would require escalation to the steering team or leadership

## Atomic Scope

Only content editing is explicitly requested in this task.

## Output Requirements

- Do not modify the source document in place.
- Save the revised document to `/root/submission/ppt_008_bullet_editing_seed.pptx`.
- Keep the output format the same as the input document (`.pptx`).
- Focus on the requested atomic operation; unrelated latent issues may remain unless they must be touched to complete the requested change.

## Expected Output Type

A revised .pptx deck with three professionally rewritten but anchor-preserving bullets.

## Why This Variant Is Hard

- The original wording is clear but not client-safe
- The revised bullets must preserve urgency without sounding hostile
- The change should stay local to the existing slide content
