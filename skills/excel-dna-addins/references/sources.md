# Source Map and Evidence Notes

## Contents

- [Skill construction](#skill-construction)
- [Excel-DNA core](#excel-dna-core)
- [Packages](#packages)
- [Key feature docs to track](#key-feature-docs-to-track)
- [Microsoft references](#microsoft-references)
- [Known stale/fragile areas to re-check before release](#known-stalefragile-areas-to-re-check-before-release)

This package should be kept aligned with the following source families.

## Skill construction

- OpenAI Codex / Agent Skills documentation: skill folder with `SKILL.md`, optional `scripts/`, `references/`, `assets/`, metadata discovery, progressive disclosure.
- OpenAI API skills documentation: versioned bundle, `SKILL.md` manifest, zip upload with one top-level folder.
- Public Jeffrey Emanuel skill pages/posts: visible pattern of deep document package with many files, workflows, templates, scorecards, and operator primers.

## Excel-DNA core

- Main docs: https://excel-dna.net/docs/
- Main repository: https://github.com/Excel-DNA/ExcelDna
- Samples repository: https://github.com/Excel-DNA/Samples
- Tutorials repository, especially RibbonBasics: https://github.com/Excel-DNA/Tutorials/tree/master/Fundamentals/RibbonBasics
- Google Group: https://groups.google.com/g/exceldna

## Packages

- `ExcelDna.AddIn`
- `ExcelDna.Testing`
- `ExcelDna.IntelliSense`
- `ExcelDna.AddIn.NativeAOT` preview line for NativeAOT paths

## Key feature docs to track

- Getting Started
- Runtime Support
- Function Registration
- Async Functions
- Excel C API
- Ribbon
- Office Ribbon Custom UI control and callback reference:
  - Part 1 overview: https://msdn.microsoft.com/en-us/library/aa338202.aspx
  - Part 2 controls and callbacks: https://msdn.microsoft.com/en-us/library/aa338199.aspx
  - Part 3 FAQ and callback signatures: https://msdn.microsoft.com/en-us/library/aa722523.aspx
- Excel 4 macro/XLM help corpus: local reference files under `C:/Work/ExcelDna/MacroHelp`, especially `TerryAney-copilot.excel.macro.reference/copilot.excel.macro.reference.md`
- Installing Add-ins
- NativeAOT support
- WiXInstaller repository

## NativeAOT preview sources to track

- Excel-DNA NativeAOT support docs and examples.
- `ExcelDna.AddIn.NativeAOT` NuGet package metadata and preview release notes.
- Excel-DNA Google Group NativeAOT threads, especially package-mixing, IntelliSense, and publish-output guidance.
- Excel-DNA GitHub issues for NativeAOT package layout/tool path conflicts and AOT warnings.
- Microsoft .NET NativeAOT deployment, analyzer, trimming, and interop documentation.
- Microsoft .NET support lifecycle for target framework recommendations.

## Microsoft references

- Office Add-ins / JavaScript custom functions overview.
- Excel XLL / C API documentation.
- .NET support lifecycle.

## Known stale/fragile areas to re-check before release

- Target framework recommendations in quickstart.
- Current supported modern .NET runtime range in Excel-DNA package/source.
- Exact `FunctionExecutionHandler` API names and signatures.
- NativeAOT feature support, package version, target framework, supported Excel-DNA preview examples, and unsupported surfaces.
- IntelliSense current stable/prerelease status.
- WiX template current toolchain.
