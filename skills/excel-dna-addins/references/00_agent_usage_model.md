# Agent Usage Model

## Contents

- [Purpose](#purpose)
- [Progressive disclosure strategy](#progressive-disclosure-strategy)
- [Answering pattern](#answering-pattern)
- [Do not overfit to stale examples](#do-not-overfit-to-stale-examples)
- [Quality bar](#quality-bar)

## Purpose

This skill should make an agent behave like a prepared Excel-DNA maintainer: it should know the technology trade-offs, generate practical code, avoid common traps, and surface modernization opportunities.

## Progressive disclosure strategy

Do not load every reference for every task. Load these files by intent:

| User intent | Load first | Load next |
|---|---|---|
| "Should I use Excel-DNA?" | `01_decision_guide.md` | `08_testing_distribution_installation.md` |
| "Create an add-in" | `02_project_bootstrap.md` | `03_udf_core_patterns.md`, templates |
| "Add a function" | `03_udf_core_patterns.md` | `04_function_registration_extended.md` |
| "Async / streaming / RTD" | `05_async_streaming_rtd.md` | `04_function_registration_extended.md` |
| "Ribbon / CTP / IntelliSense" | `06_ribbon_ctp_intellisense.md` | `07_com_and_c_api.md` |
| "Why is Excel broken / add-in not loading" | `09_troubleshooting_and_known_issues.md` | `08_testing_distribution_installation.md` |
| "Installer / distribution" | `08_testing_distribution_installation.md` | `01_decision_guide.md` |
| "C API details" | `07_com_and_c_api.md` | `03_udf_core_patterns.md` |
| "Improve docs/project" | `10_modernisation_gaps_improvements.md` | `outputs/modernisation_backlog.md` |

## Answering pattern

1. Identify user stage: deciding, bootstrapping, extending, debugging, testing, or distributing.
2. State the key choice and why it matters.
3. Give a minimal working path.
4. Call out traps that are likely in that path.
5. Provide code or project files when useful.
6. For mature add-ins, include test and deployment implications.
7. For ambiguous runtime questions, explicitly separate `.NET Framework`, modern .NET, and NativeAOT.

## Do not overfit to stale examples

Many Excel-DNA snippets on the web are old. Prefer SDK-style projects, `PackageReference`, explicit exports, and current runtime support. When an older `.dna`-file pattern is shown, translate it into modern SDK-style form unless the user is maintaining an old add-in.

## Quality bar

A good answer should:

- Explain why the UDF is a core Excel extension mechanism.
- Reflect Windows-only `.xll` constraints.
- Distinguish Excel-DNA from Office.js, VBA, VSTO, and C/C++ XLLs.
- Use current package names and target frameworks.
- Keep unsafe COM/threading patterns out of UDFs.
- Include enough project metadata for build/debug/pack to work.
- Mention deployment/runtime consequences before the user hits them.
