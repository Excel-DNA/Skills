# Excel-DNA Skills Repository

This repository develops and evaluates agent skills for Excel-DNA add-in development.

Read `AGENTS.md` first for repository layout, testing tiers, safety rules, and Git workflow. Treat `skills/excel-dna-addins/` as the canonical skill source and `evals/excel-dna-addins/` as the evaluation campaign.

## Current Baselines

- Use `net10.0-windows` as the current modern .NET baseline for controlled Windows desktop deployments.
- Use `net48` as the conservative default for broad or uncontrolled distribution.
- Treat Excel-DNA NativeAOT as preview/specialized unless current source and package evidence proves otherwise.
- Keep generated runs under `runs/`, temporary work under `scratch/`, and external source clones under `external/`.

## Claude Code Usage

- Use project slash commands from `.claude/commands/` for repeatable validation and campaign work.
- Use project subagents from `.claude/agents/` for bounded side tasks such as source verification, scenario scoring, and template smoke checks.
- Do not commit `.claude/settings.local.json` or any local credentials.
- Prefer deterministic commands over free-form evaluation claims. Do not claim a scenario passed unless the command was run and captured.

## High-Value Commands

```powershell
python evals\excel-dna-addins\scripts\run_static_skill_checks.py skills\excel-dna-addins
python evals\excel-dna-addins\scripts\run_template_smoke.py --target net10 --out scratch\template-smoke-net10
python evals\excel-dna-addins\scripts\run_claude_campaign.py --skill-dir skills\excel-dna-addins --scenarios evals\excel-dna-addins\evals\scenarios.jsonl --out runs\claude-p0-explicit --priority P0 --activation explicit --dry-run
```
