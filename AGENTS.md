# AGENTS.md

## Mission

This repository develops, tests, and iterates a deep agent skill for helping users build Microsoft Excel `.xll` add-ins with Excel-DNA and .NET.

The skill must be useful to:

- users deciding between Excel-DNA, Office.js, VBA/VSTO, and native C/C++ XLLs;
- Excel power users who are new to .NET;
- .NET developers who are new to Excel extensibility;
- existing Excel-DNA maintainers extending add-ins with UDFs, ribbons, CTPs, IntelliSense, testing, packaging, or deployment;
- advanced users working near Excel's C API, async calculation, streaming, object handles, NativeAOT, and deployment edge cases.

Treat this repo as both a skill-development project and a modernization-research project for Excel-DNA documentation, samples, tooling, templates, and user workflows.

## Preferred repository layout

Normalize the local checkout to this shape:

```text
Skills/
  AGENTS.md
  README.md
  .gitignore
  .gitattributes

  skills/
    excel-dna-addins/                  # canonical skill package source
      SKILL.md
      references/
      scripts/
      assets/
      tasks/
      evals/
      outputs/
      agents/

  .agents/
    skills/
      excel-dna-addins/                # local Codex skill activation copy or junction

  evals/
    excel-dna-addins/                  # skill testing and improvement campaign
      TEST_PLAN.md
      coverage_matrix.xlsx
      evals/
      scripts/
      rubrics/
      docs/
      reports/
      fixtures/
      ci/

  runs/                                # generated campaign outputs; gitignored
  scratch/                             # temporary scenario workspaces; gitignored
  external/                            # optional local clones of Excel-DNA repos; gitignored
  artifacts/                           # packaged zips and generated archives; gitignored by default
```

The current extracted folders may be named:

```text
excel-dna-addins
excel-dna-addins-skills-testing
```

Rename or move them to:

```text
skills/excel-dna-addins
evals/excel-dna-addins
```

For Codex skill activation, create `.agents/skills/excel-dna-addins` as either:

1. a Windows junction to `skills/excel-dna-addins`, preferred for local development; or
2. a copied mirror of `skills/excel-dna-addins`, only if junctions are inconvenient.

Suggested PowerShell/junction setup:

```powershell
New-Item -ItemType Directory -Force .agents\skills | Out-Null
cmd /c mklink /J ".agents\skills\excel-dna-addins" "skills\excel-dna-addins"
```

Do not commit generated run outputs, temporary project workspaces, NuGet caches, local source clones, or large archives unless explicitly requested.

## Source-of-truth hierarchy

When resolving factual conflicts, use this order:

1. Excel-DNA source code and build files.
2. Current Excel-DNA package metadata and release notes.
3. Current Excel-DNA documentation.
4. Excel-DNA Samples, IntelliSense, WiXInstaller, and related repositories.
5. Google Group discussions, especially for version-specific behavior, known traps, and maintainer intent.
6. Microsoft documentation for Excel, Office Add-ins, .NET, NativeAOT, COM, and the Excel C API.
7. User-provided context in the current session.

If a fact depends on a current version, package lifecycle, support policy, or preview feature status, verify it from a current source before presenting it as stable.

## Excel-DNA skill content principles

Keep the skill UDF-first. User-defined functions are the core Excel-DNA extension mechanism and the main reason many users choose `.xll` add-ins.

The skill should also cover, with accurate caveats:

- explicit exports and worksheet function metadata;
- optional/default parameters, ranges, arrays, error values, object handles, and conversions;
- async functions, streaming functions, RTD-backed patterns, and calculation semantics;
- extended registration and execution handlers as aspect-like function enhancement mechanisms;
- ribbons, custom task panes, commands, COM integration, and Excel object model access;
- ExcelDna.IntelliSense and metadata-driven function help;
- ExcelDna.Testing, build validation, integration testing, packed add-ins, signing, Trust Center issues, WiX installers, and enterprise deployment;
- .NET Framework versus modern .NET versus NativeAOT versus native C/C++ XLL trade-offs;
- correct C API concepts where Excel-DNA exposes or depends on them.

NativeAOT support is important but must be described as preview/specialized unless the current source and package state prove otherwise. Do not present NativeAOT as the default recommendation. Be precise about package separation, publish artifacts, trimming/AOT warnings, unsupported features, and Windows/Excel bitness assumptions.

## Testing philosophy

This repo should test the skill as a product, not as a prompt.

Separate these questions:

1. Does the skill trigger for the right user requests and avoid unrelated ones?
2. Does the skill route efficiently to the right reference material?
3. Is the Excel-DNA guidance factually correct and current?
4. Does generated code build?
5. Does generated code load and behave correctly in real desktop Excel?
6. Does the skill produce maintainable docs, scripts, and project structures?
7. Does a change improve general behavior rather than overfitting a single scenario?

Use explicit skill invocation for content baselines, and implicit skill invocation for trigger/routing evaluation. Keep those scores separate.

## Standard local test tiers

Run these tiers in order as the repo matures:

### Tier 0: static skill/package health

Run on every meaningful skill edit.

```powershell
python evals\excel-dna-addins\scripts\run_static_skill_checks.py skills\excel-dna-addins
```

Expected result for release candidates: zero errors and zero warnings.

### Tier 1: explicit P0 behavior

Run P0 scenarios with the skill explicitly invoked. This evaluates content without making trigger behavior the blocker.

```powershell
python evals\excel-dna-addins\scripts\run_codex_campaign.py `
  --skill-dir skills\excel-dna-addins `
  --scenarios evals\excel-dna-addins\evals\scenarios.jsonl `
  --out runs\p0-explicit `
  --priority P0 `
  --activation explicit

python evals\excel-dna-addins\scripts\score_campaign.py `
  --scenarios evals\excel-dna-addins\evals\scenarios.jsonl `
  --run-root runs\p0-explicit `
  --out runs\p0-explicit\score.json
```

### Tier 2: implicit trigger behavior

Run a smaller P0/P1 slice without explicit skill invocation.

```powershell
python evals\excel-dna-addins\scripts\run_codex_campaign.py `
  --skill-dir skills\excel-dna-addins `
  --scenarios evals\excel-dna-addins\evals\scenarios.jsonl `
  --out runs\p0-implicit-smoke `
  --priority P0 `
  --activation implicit `
  --limit 10
```

### Tier 3: buildable generated projects

For scenarios that create projects, run restore/build/validator checks. Prefer fresh scratch workspaces under `scratch/` and do not reuse contaminated directories.

### Tier 4: real Excel integration

Use a Windows machine with desktop Excel installed. Record Excel version, bitness, .NET SDKs/runtimes, NuGet state, and any Trust Center or security settings. Use ExcelDna.Testing where appropriate.

### Tier 5: cross-model/cross-agent behavior

Run a canonical subset against the agent frameworks you plan to support. Keep the scenario set identical and compare outputs using the same scoring rules.

## Scoring and failure handling

Every failed scenario should be classified as one or more of:

- trigger/routing failure;
- skill content failure;
- stale ecosystem fact;
- Excel-DNA source/docs discrepancy;
- template or script failure;
- build failure;
- real Excel runtime failure;
- unsafe COM/C API guidance;
- NativeAOT preview-status or AOT-compatibility failure;
- grader/scorer weakness;
- local infrastructure issue.

Fix the smallest general cause. Do not patch the skill only to satisfy one prompt. After a fix, rerun:

1. the failed scenario;
2. adjacent scenarios in the same topic;
3. a hold-out smoke slice;
4. static checks.

Do not claim a scenario passed unless the command was actually run and the result is captured.

## Git and GitHub workflow

Use small, reviewable commits.

Suggested branches:

```text
main
feature/nativeaot-preview-docs
feature/eval-campaign
fix/trigger-routing
fix/templates
```

Before the first GitHub push:

1. normalize the folder layout;
2. create `.gitignore` and `.gitattributes`;
3. run static checks;
4. make an initial commit;
5. create a private GitHub repo unless explicitly told to make it public.

The visible GitHub login in the connected environment is `govert`. If using GitHub CLI and no organization is explicitly selected, default to a private repo under that account, for example:

```powershell
gh repo create govert/excel-dna-agent-skills --private --source . --remote origin --push
```

If the `Excel-DNA` organization is available locally and intentionally selected, a good repo name would be:

```text
Excel-DNA/excel-dna-agent-skills
```

Do not push secrets, local run traces containing private paths or tokens, package caches, Office logs, or large generated archives.

## Suggested `.gitignore`

At minimum, ignore:

```text
runs/
scratch/
external/
artifacts/
*.zip
*.nupkg
*.snupkg
bin/
obj/
.vs/
*.user
*.suo
*.log
TestResults/
```

If `.agents/skills/excel-dna-addins` is a junction or mirror, decide whether it is committed or ignored. Prefer committing the canonical skill source under `skills/excel-dna-addins` and using a local setup step for `.agents/skills/excel-dna-addins`.

## Codex operating instructions

When working in this repo:

1. Read this file first.
2. Inspect the actual directory tree before modifying anything.
3. Preserve existing skill content unless a task requires changes.
4. Prefer PowerShell-compatible commands on Windows.
5. Keep generated outputs under `runs/` or `scratch/`.
6. Keep skill edits general, source-backed, and regression-tested.
7. When changing NativeAOT guidance, explicitly preserve preview caveats unless current sources justify changing the status.
8. When improving tests, avoid writing scenarios that only test the exact phrasing of the current skill.
9. When improving docs, update the modernization backlog if the change reveals a documentation, template, tooling, or product opportunity.
10. Summarize changes with commands run, results, files changed, and next recommended test tier.

## Useful sub-agent roles

Use sub-agents for investigation and improvement, not for official baseline scoring.

Useful roles:

- runner: executes scenario batches and captures traces;
- scorer: runs deterministic scoring and summarizes failures;
- source verifier: checks contested facts against Excel-DNA source/docs/samples and Microsoft docs;
- support miner: extracts recurring issues from Google Group history;
- skill editor: proposes minimal skill edits for a class of failures;
- template fixer: repairs scaffold/template/script issues;
- regression reviewer: checks whether a proposed fix harms adjacent scenarios or overfits.

Official scores should remain deterministic: one scenario, one fresh workspace, one agent run, one captured trace, one score.
