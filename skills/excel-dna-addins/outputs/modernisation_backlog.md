# Excel-DNA Modernisation Backlog — First Collation

## Executive view

The highest leverage work is to make the first 30 minutes reliable for new users and the first 30 days reliable for production users. That means: current runtime guidance, modern project templates, explicit UDF registration defaults, async/streaming examples that match v1.9 behavior, and deployment/trust troubleshooting.

## Immediate P0 items

1. **Refresh quickstart target frameworks**
   - Replace `.NET 6 LTS` as the highlighted modern path.
   - Present `net48` vs current supported modern .NET as a deliberate choice.
   - Explain `.NET Desktop Runtime` and roll-forward in the quickstart margin.

2. **Rewrite async guide**
   - Lead with `Task<T>` and `IObservable<T>` registration.
   - Move `ExcelAsyncUtil.Run` into a legacy/advanced section.
   - Add examples with cancellation, timeout, and error mapping.

3. **Complete Function Registration guide**
   - Type matrix.
   - Optional/default/nullables.
   - Async/streaming/object handles.
   - Handler/selector examples.
   - Explicit exports and production defaults.

4. **Add distribution/security troubleshooting page**
   - Blocked downloaded `.xll`.
   - Trust Center.
   - Code signing.
   - Disabled add-ins.
   - Runtime missing/conflict.
   - Bitness mismatch.

5. **NativeAOT preview quickstart and migration guide**
   - Make the preview status explicit.
   - Show the NativeAOT-only project shape.
   - State the exact published artifact to load.
   - List unsupported/unproven surfaces such as IntelliSense overlay, CTPs, ribbon images, rich COM assumptions, and dynamic loading.
   - Add AOT-warning troubleshooting for `IL2026`, `IL3050`, and missing native code/metadata.

## P1 items

6. **Add NativeAOT preview docs/tooling track**
   - Add top-level runtime decision docs for NativeAOT.
   - Publish a NativeAOT quickstart with `ExcelDna.AddIn.NativeAOT`, `net10.0-windows`, `win-x64`, and `PublishAot`.
   - Document supported, unsupported, and unproven features.
   - Add analyzer/build warnings for mixed package references, missing `PublishAot`, missing `win-x64`, and `ExcelDna.IntelliSense` in NativeAOT projects.
   - Add clean-machine smoke tests for the published `*-AddIn64.xll`.

7. **Create dotnet templates**
   - UDF-only net48.
   - UDF-only modern .NET.
   - Multi-target.
   - Ribbon.
   - Async/streaming.
   - Testing.

8. **Formalize COM/threading guide**
   - `ExcelDnaUtil.Application` allowed contexts.
   - `QueueAsMacro`.
   - No background COM.
   - No ordinary manual `ReleaseComObject`.

9. **Make object handles first-class docs**
   - Lifecycle.
   - Disposal.
   - Recalc/update semantics.
   - Diagnostics.
   - Sample workbook.

10. **Installer quickstart**
   - From packed `.xll` to MSI.
   - WiX/HeatWave setup.
   - Runtime prerequisite.
   - Upgrade/uninstall tests.

## P2 items

11. **Modernize samples**
   - Update READMEs.
   - Add expected output screenshots/workbook formulas.
   - Add CI build matrix.

12. **Create agent-ready docs index**
    - Feature -> canonical docs -> source files -> sample projects -> FAQ issues.

## Suggested first release package of docs improvements

Ship as one coherent refresh:

- New "Choose Excel-DNA" page.
- Updated quickstart.
- Runtime decision page.
- UDF basics page.
- Function registration page.
- Async/streaming page.
- Distribution troubleshooting page.
- One modern SDK-style sample project for each, including a NativeAOT preview template and publish validation path.

This would remove the highest-friction support questions without waiting for deeper feature work.
