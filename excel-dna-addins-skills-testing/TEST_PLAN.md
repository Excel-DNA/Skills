# Excel-DNA Agent Skill: Test, Improvement, and Validation Campaign

Generated: 2026-04-28

## 1. Purpose

This campaign validates the `excel-dna-addins` agent skill as a reusable, deep skill for helping users decide, design, build, test, migrate, troubleshoot, and distribute Excel-DNA add-ins.

The campaign treats the skill itself as a product. It tests whether the skill:
- triggers when Excel-DNA / .NET / XLL add-in work is requested;
- stays out of the way when the user really needs another technology;
- routes the agent to the right reference files without loading the whole package;
- produces technically correct Excel-DNA guidance and buildable starter code;
- protects users from high-risk bad advice around runtime choice, Excel calculation context, COM threading, C API context, and distribution security;
- discovers documentation holes and low-hanging product/usage improvements.

## 2. Grounding principles

The plan follows these skill-evaluation principles:

1. **Measure the skill as an end-to-end agent behavior.** A useful eval is `prompt -> captured run/trace/artifacts -> checks -> comparable score`.
2. **Use mixed graders.** Use code-based graders for facts that can be checked deterministically, model/rubric graders for nuanced advice, and human Excel-DNA SME review for calibration and high-risk cases.
3. **Separate capability evals from regression evals.** Capability evals should include hard scenarios and may start below target; regression evals protect the behavior already known to work.
4. **Validate progressive disclosure.** The package is intentionally encyclopedic, so efficiency depends on correct trigger metadata, a concise `SKILL.md`, and direct links to deeper references only when needed.
5. **Turn failures into skill and project improvements.** Every failure is classified as a skill instruction problem, missing reference problem, stale ecosystem fact, template/script defect, docs hole, source ambiguity, or genuine library issue.

## 3. Test objects

The campaign covers:

- `SKILL.md` metadata and routing instructions.
- Reference documents for decision, bootstrap, UDFs, registration, async/streaming, UI, COM/C API, testing, distribution, troubleshooting, and modernization.
- Templates under `assets/templates/`.
- Scripts such as `scaffold_excel_dna_addin.py` and `validate_excel_dna_project.py`.
- Any future evaluator, source-mining, or migration helper scripts added to the skill.
- Generated user artifacts: `.csproj`, `.cs`, `.dna`, `.xll`, tests, WiX files, docs, and troubleshooting plans.

## 4. Coverage dimensions

### Personas

| ID | Persona | Skill level | Primary need |
|---|---|---|---|
| P01 | Technology evaluator / solution architect | None to intermediate | Choosing between Excel-DNA, Office.js, VBA/VSTO, C/C++ XLL, or COM/VSTO for a new add-in. |
| P02 | Excel power user with beginner .NET | Beginner | Wants a working UDF add-in quickly and needs plain explanations. |
| P03 | .NET developer new to Excel | Intermediate .NET, low Excel extensibility | Understands C# and NuGet but not Excel add-ins, UDF rules, or Excel process constraints. |
| P04 | Quant / financial engineer / data scientist | Intermediate to advanced | Needs high-performance worksheet functions, arrays, async data fetches, streaming, object handles. |
| P05 | Existing Excel-DNA maintainer | Advanced | Owns legacy add-ins and needs migration, modernization, troubleshooting, and advanced registration. |
| P06 | Enterprise deployment / IT engineer | Intermediate | Needs packaging, installer, code signing, Trust Center, runtime, bitness, update strategy. |
| P07 | UI add-in developer | Intermediate | Adds ribbon, custom task panes, COM automation, IntelliSense integration. |
| P08 | Test / CI engineer | Intermediate to advanced | Needs unit tests, Excel integration tests, reproducible builds, validation in CI. |
| P09 | C API / native interop specialist | Advanced | Needs Excel C API semantics, xlf/xlc, registration, memory/threading boundaries, low-level troubleshooting. |
| P10 | Docs / contributor / support triager | Advanced | Uses skill failures to discover docs holes, FAQ gaps, and product-improvement opportunities. |

### Complexity levels

| ID | Level | Scope |
|---|---|---|
| C0 | Decision and orientation | Advice-only; no project changes. Tests technology comparisons and runtime choices. |
| C1 | Minimal add-in | Create a small add-in with one or two UDFs and metadata. |
| C2 | Production UDF layer | Typed functions, ranges, arrays, optional/default params, explicit exports, error mapping, volatility/thread-safety. |
| C3 | Advanced calculation | Async, streaming, RTD, object handles, registration transforms, execution handlers. |
| C4 | UI extensions | Ribbon, CTP, COM automation, IntelliSense, callback wiring. |
| C5 | Testing and distribution | ExcelDna.Testing, unit/integration tests, packed XLL, WiX, signing, runtime/bitness. |
| C6 | Migration and troubleshooting | SDK-style migration, package updates, runtime collisions, XLL loading issues, disabled add-ins. |
| C7 | C API / internals-aware work | Excel C API details, XLOPER/XLOPER12 semantics, registration, macro context, threading. |

### Skill levels

| ID | User skill level | Meaning |
|---|---|---|
| S0 | Newcomer | Little to no Excel-DNA background. |
| S1 | Beginner implementer | Can follow a scaffold and edit small code examples. |
| S2 | Intermediate builder | Understands C# projects and can reason about Excel calculation behavior. |
| S3 | Advanced maintainer | Can debug add-ins, migration issues, runtime boundaries, and performance trade-offs. |
| S4 | Expert contributor | Can inspect source code, docs, samples, C API behavior, and propose library/docs changes. |

## 5. Empirical goals

| ID | Metric | Goal | Measurement |
|---|---|---|---|
| G01 | Static skill package validation | 100% pass for P0 gates | run_static_skill_checks.py |
| G02 | Trigger recall on positive Excel-DNA prompts | >= 0.95 | runner trace / skill activation logs |
| G03 | Trigger specificity on negative/Office.js-only prompts | >= 0.90 | runner trace / no activation or correct deflection |
| G04 | Reference routing accuracy | >= 0.90 | trace file-read analysis + rubric |
| G05 | Reference loading efficiency | median <= 3 reference files per single-topic task; P95 <= 6 | trace file-read count |
| G06 | Domain factual correctness | >= 0.92 weighted; P0 critical errors = 0 | deterministic checks + SME-calibrated judge |
| G07 | Runtime recommendation correctness | >= 0.95; P0 critical errors = 0 | rubric + regex stale-phrase checks |
| G08 | UDF core code viability | >= 0.95 build pass for C1/C2 P0 cases | dotnet build + static checks |
| G09 | Advanced calculation code viability | >= 0.85 build/static pass; >= 0.75 Excel smoke baseline initially | build + optional self-hosted Excel |
| G10 | Excel integration pass | >= 0.85 for P0 integration cases after baseline, then ratchet | ExcelDna.Testing on Windows runner with Office |
| G11 | Distribution guidance completeness | >= 0.90 | rubric + checklist checks |
| G12 | Efficiency: answer-only tasks | No shell/build commands unless user asks; median <= 2 skill references | trace analysis |
| G13 | Efficiency: scaffold tasks | Median <= 8 shell commands and <= 4 file rewrites for minimal add-in | trace analysis |
| G14 | Regression stability | All P0 green; weighted score drop <= 1.5% versus previous accepted skill | paired A/B eval |
| G15 | Human/LLM judge agreement | >= 0.80 agreement on major-failure labels | SME sample review |

Critical failures are not averaged away. A skill version with a P0 critical failure is not accepted even if the aggregate score is high.

## 6. Scoring model

Each scenario receives:

- **Deterministic score**: regex/static checks, file existence, project build, test pass/fail, source inspection.
- **Rubric score**: 0-5 for domain correctness, user-fit, completeness, maintainability, and risk handling.
- **Efficiency score**: reference reads, commands, wall time, token/cost proxies, repeated failed attempts.
- **Critical-failure gate**: binary gate for unsafe/wrong platform advice, stale lifecycle advice, build-breaking templates, or dangerous COM/C API/security guidance.

Recommended weighting for the first campaign:

| Category | Weight |
|---|---:|
| Domain correctness | 35% |
| Artifact/build correctness | 25% |
| Process/routing correctness | 15% |
| User fit and clarity | 10% |
| Efficiency | 10% |
| Modernization/documentation signal quality | 5% |

For P0 scenarios, any critical failure sets the scenario result to fail regardless of weighted score.

## 7. Automation architecture

1. **Install candidate skill** in a clean runner workspace.
2. **Run static package checks** before invoking an agent.
3. **Run scenario prompts** with the target agent/model set.
4. **Capture traces**: final answer, files read, commands run, files written, stdout/stderr, token/cost metadata if available, and produced artifacts.
5. **Grade deterministically** where possible.
6. **Run rubric graders** for advice quality and nuanced technical correctness.
7. **Run SME review sample** for calibration and high-risk failures.
8. **Aggregate and compare** candidate skill versus previous accepted skill.
9. **Patch skill/docs/templates/scripts** based on failure taxonomy.
10. **Rerun changed families plus P0 regression suite.**

## 8. Evaluation tiers

### Tier 0: Static skill package validation

Run on every change:
- YAML frontmatter exists and has valid `name` and `description`.
- `description` front-loads Excel-DNA, Excel add-in, .xll, UDF, .NET, ribbon, CTP, IntelliSense, testing, distribution, troubleshooting trigger terms.
- `SKILL.md` body remains concise enough to act as navigation.
- All referenced files exist and are no more than one reference hop from `SKILL.md`.
- Long reference files have a table of contents.
- Script files can run `--help` or a smoke mode.
- Templates contain current package names and supported target framework guidance.
- Stale or dangerous phrases are blocked, for example `'.NET 6.0 (Long-term support)'` as a current recommendation, `XLL works on Mac`, or `disable all Excel security`.

### Tier 1: Trigger and routing validation

Positive prompts should trigger the skill. Negative prompts should not. Near-miss prompts should either deflect or route to the right neighboring technology.

Measure:
- trigger recall;
- trigger specificity;
- references loaded per scenario;
- whether the first loaded reference matches the task family.

### Tier 2: Answer-only domain QA

Run no-code prompts for decision, runtime, UDF semantics, async/streaming, UI extension, C API, testing, distribution, and troubleshooting.

Measure:
- critical fact correctness;
- balanced trade-offs;
- current lifecycle awareness;
- no unsafe advice;
- useful next steps.

### Tier 3: Code-generation and scaffolding QA

Ask the agent to create or modify add-in projects.

Deterministic checks:
- `dotnet build`;
- package and target framework checks;
- source inspection for `[ExcelFunction]`, `[ExcelArgument]`, `[ExcelHandle]`, `Task<T>`, `IObservable<T>`, `ExcelRibbon`, `CustomTaskPaneFactory`, `ExcelDna.Testing`;
- validator script pass/fail;
- no accidental exports in production scenarios.

### Tier 4: Real Excel integration QA

Run on a self-hosted Windows runner with supported Office/Excel installed:
- load generated XLL;
- verify formulas recalculate;
- verify async result appears within timeout;
- verify streaming updates at least once;
- verify ribbon callback smoke where practical;
- run ExcelDna.Testing tests;
- capture Excel logs, screenshots only when needed, and crash evidence.

### Tier 5: Distribution QA

Validate:
- packed XLL output exists;
- x86/x64 build selection is explicit;
- runtime prerequisites are documented;
- signing/Trust Center/MOTW guidance is present;
- WiX installer plan or template compiles where possible;
- clean install/uninstall or load/unload smoke tests run where possible.

### Tier 6: Modernization and documentation discovery

Every failed or confusing scenario produces one of:
- skill patch;
- docs issue;
- source/test investigation;
- sample/template improvement;
- support FAQ entry;
- low-hanging library/tooling improvement.

## 9. Scenario corpus

The complete scenario corpus is in:

- `evals/scenarios.jsonl`
- `coverage_matrix.xlsx`, sheet `Scenario Matrix`

The current corpus contains 90 scenarios across all personas and complexity levels. P0 scenarios form the release gate. P1/P2 scenarios form capability and expansion coverage.

## 10. Release gates

A candidate skill version can be accepted when:

1. Static package validation is 100% green.
2. All P0 scenarios pass or have a documented accepted exception.
3. Weighted score is at least 90% on P0+P1 regression scenarios.
4. Trigger recall is at least 95% on positive prompts.
5. Trigger specificity is at least 90% on negative prompts.
6. No critical runtime, platform, COM/threading, C API, or security-distribution failures occur.
7. Buildable starter scenarios pass at least 95% for C1/C2 P0 cases.
8. Human review agrees that rubric graders are not masking major errors.

## 11. Improvement loop

Use this loop for every iteration:

1. **Baseline** current accepted skill on P0 + representative P1.
2. **Change one thing at a time** where practical: description, SKILL.md routing, one reference file, one template, or one script.
3. **Run targeted family evals** for the changed area.
4. **Run P0 regression** before accepting the change.
5. **Classify failures** with the taxonomy below.
6. **Patch the most upstream cause**:
   - description for trigger errors;
   - SKILL.md navigation for missed references;
   - reference content for factual gaps;
   - templates/scripts for repeatable implementation errors;
   - source/docs backlog for library/documentation holes.
7. **Record the before/after result** in `reports/`.

## 12. Failure taxonomy

| ID | Failure | Meaning | Typical remediation |
|---|---|---|---|
| F01 | Trigger false negative | Skill not selected for an Excel-DNA add-in task. | Rewrite description with front-loaded trigger terms; add trigger eval variants. |
| F02 | Trigger false positive | Skill selected for unrelated spreadsheet/Python/Office.js-only task. | Narrow description or add negative routing guidance. |
| F03 | Source factual error | Answer contradicts Excel-DNA source/docs/current platform facts. | Patch reference; add deterministic regression. |
| F04 | Stale ecosystem fact | .NET support lifecycle, package names, or supported targets are outdated. | Add freshness sentinel and source-date review. |
| F05 | Unsafe Excel/COM/C API guidance | Advice risks crashes, deadlocks, UI mutation during calculation, or security bypass. | Escalate to P0; add red flag in SKILL.md/reference. |
| F06 | Build failure | Generated project does not compile. | Patch template/scaffold script; add compiler check. |
| F07 | Excel runtime failure | Build passes but XLL does not load or function fails in Excel. | Add ExcelDna.Testing smoke and troubleshooting note. |
| F08 | Overly vague answer | No concrete steps/files/code where expected. | Add recipe/template/checklist. |
| F09 | Over-specific brittle output | Agent follows one path when user constraints require another. | Move guidance from low-freedom script to medium-freedom recipe. |
| F10 | Excessive context/tool use | Agent reads too many references or runs unnecessary commands. | Improve SKILL.md navigation and reference headings. |
| F11 | Docs/source gap | Skill failure reveals docs are ambiguous or missing. | File docs backlog item and add temporary skill note. |
| F12 | Template drift | Starter templates lag package/runtime recommendations. | Version template tests and refresh lifecycle notes. |

## 13. Model and framework matrix

Run the same scenario set against each target environment:

- GPT 5.5 Pro / Codex-style coding agent with the skill installed.
- Claude Opus 4.7 / Claude Code-style coding agent with the skill installed.
- A faster/economical model profile for regression-smoke routing and simple scaffolds.
- Optional no-skill baseline to quantify skill lift.

Track:
- pass rate;
- critical failures;
- token/reference efficiency;
- time-to-artifact;
- build/test pass rate;
- evaluator agreement.

## 14. Windows/Excel integration environment

Use a self-hosted Windows runner for Excel-dependent tests. Hosted CI images usually do not provide a licensed interactive Excel installation suitable for XLL loading tests.

Recommended runner inventory:
- Windows 11 or supported Windows Server desktop-capable environment.
- Office/Excel current channel plus at least one semi-annual or legacy channel if enterprise customers require it.
- 64-bit Excel primary; 32-bit Excel runner if supporting x86.
- .NET Framework 4.8 and currently supported modern .NET Desktop Runtime(s).
- Visual Studio Build Tools or .NET SDKs.
- Code signing test certificate where needed.
- Excel configured for deterministic test startup, with isolated profile and no unrelated add-ins.

## 15. Data schema for run capture

Each run should write:

```json
{
  "scenario_id": "EDNA-0001",
  "skill_version": "0.1.0-candidate",
  "model": "gpt-5.5-pro",
  "runner": "codex-cli",
  "start_time_utc": "...",
  "end_time_utc": "...",
  "skill_triggered": true,
  "references_read": ["SKILL.md", "references/02_project_bootstrap.md"],
  "commands": ["python scripts/scaffold_excel_dna_addin.py ...", "dotnet build"],
  "files_created": ["MyAddin.csproj", "Functions.cs"],
  "deterministic_checks": {"dotnet_build": "pass"},
  "rubric_scores": {"domain_correctness": 5, "risk_handling": 5},
  "critical_failures": [],
  "failure_tags": [],
  "notes": ""
}
```

## 16. Documentation-hole and low-hanging improvement outputs

Every campaign run should produce:
- top 10 failing scenario families;
- top 10 missing or stale facts;
- top 10 template/script fixes;
- top 10 docs gaps;
- low-hanging improvements ranked by expected eval lift and implementation effort.

Examples of expected early signals:
- stale runtime guidance and .NET support lifecycle wording;
- insufficient async/streaming examples under v1.9 registration;
- COM/threading guidance buried in samples rather than core docs;
- need for a canonical distribution/troubleshooting matrix;
- need for `dotnet new` templates or analyzer warnings around accidental exports and runtime choices.
