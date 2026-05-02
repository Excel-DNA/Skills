# Testing, Packing, Signing, Distribution, and Installation

## Contents

- [Testing layers](#testing-layers)
- [ExcelDna.Testing](#exceldnatesting)
- [Packing](#packing)
- [Code signing and Trust Center](#code-signing-and-trust-center)
- [Runtime prerequisites](#runtime-prerequisites)
  - [.NET Framework](#net-framework)
  - [Modern .NET](#modern-net)
  - [NativeAOT](#nativeaot)
- [Installation modes](#installation-modes)
- [WiX installer template](#wix-installer-template)
- [Versioning](#versioning)
- [NativeAOT preview release gates](#nativeaot-preview-release-gates)
- [Release checklist](#release-checklist)

## Testing layers

Use multiple test layers because Excel-DNA add-ins live across .NET code, Excel registration, workbook behavior, and installer/runtime state.

| Layer | Purpose | Tools |
|---|---|---|
| Pure .NET unit tests | Validate domain logic without Excel | xUnit/NUnit/MSTest |
| UDF wrapper tests | Validate Excel-facing conversion/error behavior | Unit tests plus Excel-DNA integration helpers |
| Excel-hosted tests | Validate behavior inside Excel process | ExcelDna.Testing / xUnit hosting |
| Workbook smoke tests | Validate real formulas and recalc | Sample workbooks, CI/manual smoke |
| Packaging tests | Validate packed `.xll`, dependencies, signing | build scripts, local clean VM |
| Installer tests | Validate registry/install/runtime prerequisites | WiX/MSI tests, clean machine |

## ExcelDna.Testing

ExcelDna.Testing supports automatic tests for Excel models/add-ins, including Excel-DNA and VBA, hosted by xUnit. Use it when the behavior depends on Excel itself: registration, formula evaluation, recalculation, ranges, or workbook interaction.

Suggested test layout:

```
src/MyAddIn/MyAddIn.csproj
tests/MyAddIn.Tests/MyAddIn.Tests.csproj
tests/MyAddIn.ExcelTests/MyAddIn.ExcelTests.csproj
samples/SmokeWorkbook.xlsx
```

Test categories:

- Function appears in Excel.
- Function returns correct scalar values.
- Function returns correct 2D arrays.
- Function maps invalid inputs to expected Excel errors.
- Async function eventually completes.
- Streaming function updates and unsubscribes.
- Ribbon loads and command can run a no-op/safe action.
- Packed `.xll` loads on a clean machine.

## Packing

Excel-DNA can pack add-ins into `.xll` outputs for distribution. Packing reduces deployment complexity by embedding managed assemblies and resources into the add-in artifact.

Checklist:

- Build Release configuration.
- Produce both x86 and x64 only if both are required.
- Confirm dependencies are included or deliberately external.
- Load the packed `.xll` on a clean machine.
- Confirm no debug-only paths or local files are required.
- Sign the packed `.xll` for distribution.

## Code signing and Trust Center

For broad distribution, sign the packed `.xll`. Unsigned add-ins can trigger warnings or be blocked depending on Office Trust Center settings, file origin, and enterprise policy.

Common blockers:

- File downloaded from the internet and marked with Mark-of-the-Web.
- Add-in stored outside trusted locations.
- Macro/add-in policy disabled.
- Unsigned or untrusted publisher.
- Excel disabled the add-in after a load failure.

## Runtime prerequisites

### .NET Framework

For broad distribution, `.NET Framework 4.8` tends to be the conservative path. It avoids modern runtime collision issues and is widely present in Windows environments.

### Modern .NET

Modern .NET add-ins need the matching .NET Desktop Runtime installed for the Excel bitness/architecture and target framework policy. Runtime selection can be influenced by `RollForward`, but only one modern .NET runtime can be loaded into the Excel process. This affects coexistence with other add-ins.

Installer guidance:

- Detect required runtime.
- Install or prompt for .NET Desktop Runtime where policy allows.
- Choose a roll-forward policy deliberately.
- Test coexistence with other modern .NET Excel-DNA add-ins.

### NativeAOT

NativeAOT can remove the runtime-install prerequisite for compatible 64-bit add-ins, but it has feature constraints and is currently a preview/specialized path. Treat it as its own project track with separate publish, warning, clean-machine, and feature-compatibility gates.

Distribution implications:

- Publish, do not just build. Copy the published `*-AddIn64.xll` from the RID-specific publish folder.
- Test on 64-bit Excel. Current preview guidance is 64-bit-first.
- A .NET Desktop Runtime installation should not be required for the NativeAOT add-in, but Trust Center, bitness, signing, and native dependency issues still apply.
- Sign the published XLL. Runtime-free does not mean security-policy-free.
- Record the preview package version in release notes.

## Installation modes

One-time/manual:

- Open `.xll` directly from Excel.
- Use Excel Add-ins dialog.
- Use VBA `RegisterXLL` for a workbook-local load.

Persistent/user install:

- Register add-in path in Excel Add-ins collection.
- Use HKCU Excel `OPEN` registry entries.
- Use an installer to copy files to a stable location and configure Excel.

Enterprise install:

- MSI or deployment tooling.
- Trusted locations/publisher configuration.
- Runtime prerequisite management.
- Versioned install path and upgrade strategy.

## WiX installer template

Excel-DNA provides a WiX installer template repository. Use it when the add-in needs a proper Windows installer with upgrade codes, product metadata, registry entries, and deployment assets.

WiX checklist:

- Choose stable `ProductCode`/`UpgradeCode` strategy.
- Define per-user vs per-machine install.
- Register Excel add-in for the current user unless enterprise policy dictates otherwise.
- Add .NET runtime prerequisite detection for modern .NET.
- Install both bitness variants only when needed.
- Sign MSI and `.xll` artifacts.
- Test install, repair, upgrade, and uninstall.

## Versioning

Recommended artifacts:

- `MyAddIn.xll`
- `MyAddIn64.xll` if separate 64-bit naming is used
- `MyAddIn-<version>.msi`
- `CHANGELOG.md`
- sample workbook matching the release

Keep function names stable across versions. Breaking UDF signature changes break workbooks.


## NativeAOT preview release gates

A NativeAOT preview add-in should have a stricter release gate than an ordinary managed add-in:

1. `dotnet publish -c Release -r win-x64` succeeds.
2. AOT/trimming warnings are zero or every warning is documented with a targeted test.
3. The published `*-AddIn64.xll` loads in 64-bit Excel on a clean machine.
4. The machine does not need the .NET Desktop Runtime for the add-in to load.
5. Core UDF formulas calculate correctly.
6. Async/streaming/object-handle/ribbon features are tested if present.
7. `ExcelDna.IntelliSense` and CTP assumptions are absent unless current preview tests prove support.
8. The artifact is signed and tested with realistic Trust Center policy.
9. Native dependencies, if any, are available to the XLL or statically linked according to the project design.
10. A managed fallback build exists for users who cannot accept preview risk.

## Release checklist

- All public functions reviewed for name/metadata/category.
- All obsolete functions remain available or have documented migration.
- Build produces expected x86/x64 artifacts.
- Packed `.xll` loads on clean machine.
- Runtime prerequisites documented and tested.
- Code signing complete.
- Installer upgrade/uninstall tested.
- Sample workbook recalculates correctly.
- Excel disabled-addins state checked after failures.
- Smoke tests cover async/streaming if present.
