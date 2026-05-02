# Iteration and Improvement Playbook

## Triage order

1. Fix P0 critical correctness failures.
2. Fix trigger false negatives/false positives by editing the description first.
3. Fix wrong reference routing by editing SKILL.md navigation.
4. Fix repeated code-generation failures in templates or scaffold scripts.
5. Fix nuanced domain misses in reference files.
6. Convert repeated support/troubleshooting failures into documentation backlog items.
7. Add regression scenarios for every fixed failure.

## Patch selection

Prefer the smallest durable patch:
- Description change for discovery problems.
- SKILL.md routing change for navigation problems.
- Reference addition for missing domain knowledge.
- Template/script change for repeatable artifact generation.
- Validator/analyzer check for recurring project mistakes.
- Docs backlog item when the skill is compensating for a public documentation hole.

## A/B acceptance

Compare candidate skill B against accepted skill A on the same scenarios. Accept B only if:
- all P0 cases remain green;
- no new critical failures;
- aggregate weighted score improves or stays within 1.5% with a justified maintainability improvement;
- efficiency does not degrade materially unless correctness improves on critical cases.

## Human calibration

Every release sample should include:
- five passed P0 cases;
- five failed or borderline cases;
- five high-risk cases from runtime, COM/C API, and distribution;
- five random P1 cases.

Update rubric examples when human/LLM judge disagreement reveals ambiguous scoring.
