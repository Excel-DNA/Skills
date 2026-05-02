# Documentation Holes and Low-hanging Modernisation Ideas

## Contents

- [Priority documentation holes](#priority-documentation-holes)
- [Low-hanging usage improvements](#low-hanging-usage-improvements)
  - [1. `dotnet new exceldna` templates](#1-dotnet-new-exceldna-templates)
  - [2. Project analyzer / build warnings](#2-project-analyzer-build-warnings)
  - [3. Runtime chooser in docs](#3-runtime-chooser-in-docs)
  - [4. Sample workbook for each feature](#4-sample-workbook-for-each-feature)
  - [5. Installer quickstart](#5-installer-quickstart)
  - [6. Migration notes for old docs/projects](#6-migration-notes-for-old-docsprojects)
  - [7. Agent-friendly source map](#7-agent-friendly-source-map)
  - [8. NativeAOT preview tooling and docs](#8-nativeaot-preview-tooling-and-docs)
- [Documentation structure proposal](#documentation-structure-proposal)
- [Agent-skill improvements still needed](#agent-skill-improvements-still-needed)

This file is an explicit output of the collation. It separates documentation gaps from product/tooling improvements.

## Priority documentation holes

| Priority | Gap | Why it matters | Suggested fix |
|---|---|---|---|
| P0 | Getting Started still centers `.NET 6 LTS` in places | .NET 6 is out of support; new users may start on an obsolete runtime | Update quickstart to `net8.0-windows` or current supported LTS guidance, while explaining when `net48` is the safer default |
| P0 | Runtime decision guidance is spread out | Runtime choice determines deployment success and add-in coexistence | Create a one-page decision tree: `net48`, `net8+`, NativeAOT, multi-target |
| P0 | Async docs mix older Registration-library patterns with newer v1.9 registration | Users may cargo-cult stale setup or miss direct `Task<T>` / `IObservable<T>` support | Rewrite async page around modern task/observable patterns, with legacy section below |
| P0 | Function Registration page says key areas are still to be expanded | Registration is central to UDF development | Expand into type matrix, optional/default params, async/streaming, object handles, handlers, examples |
| P1 | Explicit export behavior and accidental function registration need stronger guidance | Public helper methods can accidentally become worksheet functions | Make `ExcelAddInExplicitExports=true` the production default and explain prototype exception |
| P1 | Object handle lifecycle is sample-driven | Handles are powerful but easy to leak or make opaque | Add object handle guide: creation, update, recalc, disposal, diagnostics, examples |
| P1 | CTP guidance is mostly sample-based | Task panes are an important extension surface | Add formal CTP guide: manager pattern, workbook/window scoping, disposal, thread rules |
| P1 | COM threading guidance is buried in sample notes | Wrong COM use causes hard-to-debug crashes/deadlocks | Add a central page: `ExcelDnaUtil.Application`, main thread, QueueAsMacro, no ReleaseComObject folklore |
| P1 | Distribution docs need a current trust/security section | Office blocks and Mark-of-the-Web are frequent support problems | Create troubleshooting matrix for blocked XLL, disabled add-ins, signing, trusted locations |
| P1 | WiX template needs a user journey | Installer creation is a big hurdle | Add tutorial from add-in output to MSI, including runtime prerequisite and upgrade codes |
| P2 | IntelliSense package/version status is unclear for new users | Users may not know whether to add it as separate package, extension, or embedded feature | Provide a current compatibility matrix and minimal install example |
| P2 | Testing tutorial is split from main docs | Testing is a key modernization theme | Integrate ExcelDna.Testing guide with sample test projects and CI advice |
| P0 | NativeAOT support is new, important, and easy to overgeneralize | Users may assume all features work or mix packages incorrectly | Add preview-status banner, feature compatibility table, package-layout rules, migration checklist, and clean-machine publish test path |
| P2 | Samples target mixed eras/styles | Samples teach by copy/paste | Modernize sample READMEs with SDK-style, target framework, package version, expected output |

## Low-hanging usage improvements

### 1. `dotnet new exceldna` templates

Create templates:

- `exceldna-udf-net48`
- `exceldna-udf-net8`
- `exceldna-ribbon`
- `exceldna-async`
- `exceldna-testing`
- `exceldna-wix`
- `exceldna-nativeaot-preview`

Each should include a smoke function and README that says where the `.xll` is produced.

### 2. Project analyzer / build warnings

Add warnings for:

- package `Excel-DNA` instead of `ExcelDna.AddIn`
- `net6.0-windows` / `net7.0-windows` targets
- modern .NET target without runtime/roll-forward guidance
- public static functions with implicit registration in production mode
- `IsThreadSafe=true` plus suspicious COM references
- missing function descriptions
- functions returning unsupported types
- NativeAOT project missing `PublishAot` or `RuntimeIdentifier=win-x64`
- NativeAOT project mixing `ExcelDna.AddIn` and `ExcelDna.AddIn.NativeAOT` unconditionally
- NativeAOT project referencing `ExcelDna.IntelliSense` before support is proven
- NativeAOT project with AOT/trimming warnings in publish logs

The included `scripts/validate_excel_dna_project.py` is a seed for this idea.

### 3. Runtime chooser in docs

A single chart should answer:

- Do users control target machines?
- Need new C#/.NET libraries?
- Need to coexist with other modern .NET add-ins?
- Need no runtime installation?
- Need x86?
- Need runtime-free distribution?
- Is the add-in NativeAOT-compatible?
- Need widest compatibility?

### 4. Sample workbook for each feature

For every main feature page, include a workbook:

- UDF basics
- async task
- streaming observable
- object handles
- array functions
- ribbon command
- CTP
- IntelliSense
- NativeAOT preview UDF

### 5. Installer quickstart

Give a copy/paste path:

1. Create add-in.
2. Pack `.xll`.
3. Sign `.xll`.
4. Create WiX project from template.
5. Add registry entry for Excel add-in.
6. Add runtime prerequisite.
7. Build MSI.
8. Test install/uninstall.

### 6. Migration notes for old docs/projects

Add page: "If your project has an old `.dna` file / old NuGet package / non-SDK `.csproj`, do this." Include before/after project files.

### 7. Agent-friendly source map

Maintain a machine-readable `docs/source-map.json` mapping features to canonical docs, samples, packages, source files, and known issues. This skill package includes a human-readable `references/sources.md`; a JSON version should be added later.


### 8. NativeAOT preview tooling and docs

Low-hanging NativeAOT improvements:

- Add a docs banner: preview, 64-bit-first, runtime-free, feature-limited.
- Add a copy/paste NativeAOT quickstart with current preview package version, publish command, and exact output artifact.
- Add a compatibility table covering UDFs, async, object handles, execution handlers, ribbon, ribbon images, CTPs, IntelliSense, COM access, and native dependencies.
- Add a migration checklist from ordinary managed Excel-DNA to NativeAOT.
- Add a troubleshooting table for mixed packages, wrong artifact, missing `win-x64`, AOT warnings, and IntelliSense/WinForms assumptions.
- Add a CI recipe that runs `dotnet publish -c Release -r win-x64` and fails on AOT/trimming warnings.
- Add a small benchmark workbook comparing managed modern .NET and NativeAOT for cold load, first calculation, repeated calculation, and memory footprint.

## Documentation structure proposal

Recommended doc IA:

1. Choose Excel-DNA
2. Quickstart
3. UDFs
4. Registration and metadata
5. Types and ranges
6. Async and streaming
7. Object handles
8. Ribbon and commands
9. Custom task panes
10. IntelliSense
11. COM and C API
12. Testing
13. Packaging and distribution
14. Installers
15. Troubleshooting
16. Migration and modernization
17. Reference/API compatibility

## Agent-skill improvements still needed

- Verify every code snippet against the current source tree.
- Add real compileable `.csproj` templates rather than text snippets only.
- Add source-code anchors from Excel-DNA repositories.
- Add Google Group-derived FAQ after mining recurring support questions.
- Add a C API sub-skill later with `XLOPER12`, `xlfRegister`, memory, callbacks, and type strings.
