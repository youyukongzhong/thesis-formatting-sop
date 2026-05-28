---
name: thesis-format-system-v1
description: Use when the user asks Codex to format, standardize, rescue, or final-check a thesis, dissertation, graduation paper, or other long academic Word document against a school format-requirements file, template, or teacher feedback. This skill enforces the default order: read requirements first, work on a docx copy, build styles before line-by-line edits, repair global structure before paragraph cleanup, then handle figures, tables, references, citations, and final verification.
---

# Thesis Format System V1

Use this skill to run thesis-format work as a stable system instead of ad hoc paragraph tweaking.

## Non-negotiable rules

- Work on a duplicate, never the only original.
- Prefer `docx`; if the source is `doc`, convert first and treat `docx` as the working file.
- Read the format-requirements file before editing the thesis.
- Go from large to small and from coarse to fine.
- Build or normalize the style library before broad paragraph edits.
- Repair structure before cosmetic cleanup.
- Treat figures, tables, references, and in-text citations as dedicated late-stage checkpoints.
- Do not fabricate references, data, or case facts.

## Execution order

1. Read the requirements and extract hard constraints.
2. Inspect the thesis and create a problem list.
3. Build or normalize styles.
4. Repair sections, headers, footers, page numbering, heading hierarchy, navigation, and TOC.
5. Normalize body paragraphs.
6. Handle figures and tables.
7. Handle references and in-text citations.
8. Run final verification before delivery.

## Read these references

- Read `references/flow.md` at the start of substantial thesis-format work.
- Read `references/checklist.md` before declaring the task complete.
- Read `references/guardrails.md` when the request drifts from formatting into content judgment, invented references, fake data, or teacher-opinion conflicts.

## Working preference

Use direct paragraph-by-paragraph edits only when:

- a few local paragraphs resist style application
- the document is already partially corrupted
- blank paragraphs, broken wraps, or citation superscripts need surgical cleanup

Treat `style library first` as the default path.
