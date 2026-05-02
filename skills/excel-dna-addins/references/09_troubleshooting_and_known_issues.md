# Troubleshooting and Known Issues

## Contents

- [Add-in does not load](#add-in-does-not-load)
- [Functions do not appear](#functions-do-not-appear)
- [Function returns `#VALUE!` unexpectedly](#function-returns-value-unexpectedly)
- [Ribbon does not show](#ribbon-does-not-show)
- [CTP does not show or leaks](#ctp-does-not-show-or-leaks)
- [Modern .NET add-in fails while .NET Framework add-in works](#modern-net-add-in-fails-while-net-framework-add-in-works)
- [Async function never completes](#async-function-never-completes)
- [Streaming function overwhelms Excel](#streaming-function-overwhelms-excel)
- [COM object model errors](#com-object-model-errors)
- [Packed add-in works on developer machine but not user machine](#packed-add-in-works-on-developer-machine-but-not-user-machine)
- [NativeAOT preview failures](#nativeaot-preview-failures)

## Add-in does not load

Check in this order:

1. Excel bitness matches `.xll` bitness.
2. The file is not blocked by Windows Mark-of-the-Web.
3. Office Trust Center policy allows the add-in.
4. The add-in is not in Excel's disabled add-ins list.
5. Required .NET runtime is installed.
6. Dependencies are packed or present beside the add-in.
7. Native dependencies are the right architecture.
8. The add-in path is accessible and not blocked by network policy.
9. Try loading from a simple local trusted path.

## Functions do not appear

Likely causes:

- `ExcelAddInExplicitExports=true` but functions lack `[ExcelFunction]`.
- Function method is not `public static`.
- Class/method visibility is wrong.
- Unsupported parameter/return type without registration conversion.
- Name collision with another function/add-in.
- Add-in load failed before registration.
- Old `.dna` file excludes the assembly or type.

## Function returns `#VALUE!` unexpectedly

Likely causes:

- Parameter conversion failed.
- Function threw an exception.
- Range shape differs from expected.
- Missing/empty values not handled.
- Function returns unsupported object type.
- Async/streaming registration not configured as expected.

Debug path:

1. Reproduce with the simplest formula.
2. Add error mapping/logging.
3. Test the .NET method directly.
4. Test Excel-facing wrapper with representative `object[,]` values.
5. Confirm registration metadata.

## Ribbon does not show

Likely causes:

- Ribbon class not public.
- Ribbon class does not derive directly from `ExcelRibbon`.
- `[ComVisible(true)]` missing where required.
- Invalid CustomUI XML.
- Wrong namespace.
- Callback signature mismatch.
- Excel has disabled the helper COM add-in.
- UI errors hidden.

## CTP does not show or leaks

Likely causes:

- Pane manager loses references or retains too many references.
- Pane is created for wrong window/workbook context.
- UI code runs from wrong thread.
- Content control depends on missing WinForms/WPF runtime support.
- Pane is not disposed on workbook/window close.

## Modern .NET add-in fails while .NET Framework add-in works

Likely causes:

- .NET Desktop Runtime missing.
- Runtime bitness/architecture mismatch.
- Another add-in already loaded an incompatible modern .NET runtime into Excel.
- `RollForward` policy too restrictive.
- Dependency requires a newer runtime.
- Self-contained assumption is wrong for ordinary Excel-DNA managed add-ins.

## Async function never completes

Likely causes:

- Function identity arguments omit one or more inputs.
- Async delegate deadlocks on synchronization context.
- Operation throws and error is swallowed.
- RTD updates disabled or delayed.
- Cancellation occurs when formula is removed/recalculated.
- External service call hangs without timeout.

Patterns:

- Use `ConfigureAwait(false)` in library awaits.
- Add timeouts and cancellation tokens.
- Log start/completion/error.
- Return explicit status for external service failures.

## Streaming function overwhelms Excel

Likely causes:

- Source updates too frequently.
- Lossless mode used where latest-value semantics suffice.
- No throttling/debouncing.
- Too many formulas subscribe independently instead of sharing source subscription.

Fixes:

- Throttle updates.
- Share subscriptions.
- Batch updates.
- Use latest-value streams where possible.

## COM object model errors

Likely causes:

- COM called from background thread.
- COM called from thread-safe UDF.
- COM proxy retained too long.
- Attempted workbook mutation during calculation.
- Manual `ReleaseComObject` broke a still-used object.

Fixes:

- Use ribbon callbacks/commands/queued macros.
- Keep COM interactions short and localized.
- Avoid manual COM release in ordinary code.

## Packed add-in works on developer machine but not user machine

Likely causes:

- Dependency missing because it was not packed.
- Native dependency missing or wrong architecture.
- Required runtime missing.
- Trust Center/MOTW blocks.
- Environment-specific file path or configuration.
- User has different Excel bitness/version.
- Antivirus or enterprise policy blocks XLL.

Debug with a clean VM matching the user's Excel, Windows, and policy environment.


## NativeAOT preview failures

NativeAOT preview projects have distinct failure modes. Check these before treating the problem as an ordinary managed Excel-DNA load issue.

Likely causes:

- Loaded the ordinary build output instead of the published `*-AddIn64.xll`.
- Used 32-bit Excel.
- Mixed `ExcelDna.AddIn` and `ExcelDna.AddIn.NativeAOT` unconditionally in the same project.
- Forgot `RuntimeIdentifier=win-x64` or `PublishAot=true`.
- Included `ExcelDna.IntelliSense` or a CTP/rich UI package that is not AOT-compatible.
- Ignored AOT/trimming warnings during publish.
- Dependency relies on dynamic loading, runtime code generation, broad reflection, C++/CLI, or unsupported COM.
- Preview package layout/tool-path issue such as `ExcelDnaToolsPath` resolving to the wrong package tools folder.

Debug path:

1. Reproduce with the minimal NativeAOT-only project from `references/11_native_aot_preview.md`.
2. Confirm `dotnet publish -c Release -r win-x64` is used.
3. Load the published `*-AddIn64.xll` in 64-bit Excel.
4. Remove `ExcelDna.AddIn`, `ExcelDna.IntelliSense`, CTP/UI packages, and all nonessential dependencies.
5. Add functions/features back one at a time.
6. Fix AOT warnings before testing in Excel.
7. If package path errors mention missing base XLL assets, check the current preview package issue/workaround and avoid mixed package references where possible.
