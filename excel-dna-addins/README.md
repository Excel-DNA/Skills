# Excel-DNA Add-ins Agent Skill Package — Draft 0.1

This package starts a deep, agent-oriented skill for making Excel add-ins with Excel-DNA. It is deliberately more than a single prompt: it contains a manifest, routing instructions, detailed references, templates, scripts, evaluation scenarios, and a modernization backlog. It now includes a dedicated NativeAOT preview track for runtime-free 64-bit Excel-DNA add-ins.

## Contents

- `SKILL.md` — entry point and activation/routing rules.
- `references/` — encyclopedia-style guidance.
- `assets/templates/` — starter project skeletons and code snippets, including a NativeAOT preview template.
- `scripts/` — helper scripts for scaffolding and project validation.
- `tasks/` — recipes for common work.
- `evals/` — scenarios for testing whether an agent uses the skill well.
- `outputs/` — explicit collation outputs, including documentation holes and low-hanging product improvements.

## How to use in an agentic coding framework

Place the `excel-dna-addins` directory where the framework expects skills. For OpenAI-compatible skill systems, keep the directory zipped with a single top-level folder or expose the folder directly. The `SKILL.md` frontmatter is the discovery surface; the references are intentionally separated so the model can load only the relevant sections.

## Package maturity

This is a first collation. The highest-value next step is to align the references with the current source code, convert examples into compilable sample projects with CI checks, and keep the NativeAOT preview guidance synchronized with package/source changes.
