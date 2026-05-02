# COM and Excel C API for Excel-DNA Add-ins

## Contents

- [Why C API knowledge matters](#why-c-api-knowledge-matters)
- [XLL lifecycle concepts](#xll-lifecycle-concepts)
- [`XlCall`](#xlcall)
- [`ExcelReference`](#excelreference)
- [`Excel`, `TryExcel`, and errors](#excel-tryexcel-and-errors)
- [Context categories](#context-categories)
- [COM object model access](#com-object-model-access)
- [`QueueAsMacro`](#queueasmacro)
- [Registration and C API implications](#registration-and-c-api-implications)
- [C API safety checklist](#c-api-safety-checklist)
- [NativeAOT interop implications](#nativeaot-interop-implications)
- [Separate future C API skill](#separate-future-c-api-skill)

## Why C API knowledge matters

Excel-DNA hides much of the C/XLL complexity, but correct add-in behavior still depends on Excel's native rules. UDF registration, allowed callbacks, thread safety, range references, macro contexts, and memory marshaling all trace back to Excel's C API and XLL model.

## XLL lifecycle concepts

A native XLL is a DLL loaded by Excel with exported callback procedures such as `xlAutoOpen`, `xlAutoClose`, and registration routines. Functions and commands are registered with Excel through C API calls such as `xlfRegister`.

Excel-DNA provides the loader and managed bridge so the add-in author can write managed functions while Excel sees an XLL.

## `XlCall`

`XlCall` exposes direct calls into Excel's C API from .NET.

Conceptual use:

```csharp
object result = XlCall.Excel(XlCall.xlfGetCell, 48, new ExcelReference(0, 0, 0, 0, "Sheet1"));
```

Use `XlCall` when normal Excel-DNA abstractions are insufficient. Avoid it for ordinary UDFs where typed parameters and return values are enough.

## `ExcelReference`

`ExcelReference` represents an Excel range reference. It is useful when a function needs the reference identity, dimensions, sheet, or direct C API access rather than just the values.

Do not use `ExcelReference` merely to read cell values; use typed values or `object[,]` unless reference identity matters.

## `Excel`, `TryExcel`, and errors

C API calls can fail depending on context. Prefer safe patterns that detect failure and return deliberate Excel errors rather than letting exceptions or invalid values leak.

Use cases for `TryExcel`-style calls:

- A C API call may be disallowed in the current context.
- A macro-sheet function is only conditionally available.
- You are probing Excel state for diagnostics.

## Context categories

Excel C API calls have context restrictions. Some are safe during worksheet calculation; others require macro context; some are commands or special callbacks.

Design implications:

- Do not mutate Excel state from UDF calculation.
- Do not call UI/COM APIs from thread-safe UDFs.
- Use ribbon callbacks/macros for actions.
- Use `QueueAsMacro` to transition back to a safe Excel context for state-changing operations.

## COM object model access

Use `ExcelDnaUtil.Application` to access Excel's COM object model in appropriate contexts such as ribbon callbacks, commands, and queued macros.

```csharp
public static void WriteMessageToActiveCell()
{
    ExcelAsyncUtil.QueueAsMacro(() =>
    {
        dynamic excel = ExcelDnaUtil.Application;
        excel.ActiveCell.Value2 = "Updated from macro context";
    });
}
```

Rules:

- Access Excel COM only from the main Excel thread / valid callback context.
- Do not call Excel COM APIs from background worker threads.
- Do not call Excel COM APIs from thread-safe worksheet functions.
- Avoid manual `Marshal.ReleaseComObject` / `FinalReleaseComObject` in ordinary Excel-DNA code; it can break objects still owned by Excel or other code paths.
- Avoid retaining COM proxies longer than needed.

## `QueueAsMacro`

`ExcelAsyncUtil.QueueAsMacro` schedules code to run in a macro-capable context. Use it for:

- updating cells after async completion outside a UDF return path
- showing dialogs from non-UI code
- workbook state changes initiated indirectly
- bridging from background work to Excel main thread

Keep queued macros short and robust.

## Registration and C API implications

Excel-DNA handles registration for ordinary functions. But when troubleshooting function registration, remember:

- Excel registers names and type signatures, not C# methods directly.
- Function metadata maps to Excel registration data.
- Thread-safe and macro-type flags affect how Excel may call the function.
- Some type combinations require conversion/wrapping.
- Function names are global within the Excel session and can collide.

## C API safety checklist

Before adding direct `XlCall` usage:

- Is a normal Excel-DNA parameter/return type enough?
- Is the call allowed during UDF calculation?
- Does the function need `IsMacroType`?
- Is the function thread-safe if marked so?
- What Excel error should be returned if the call fails?
- Is the call version-specific?
- Does it depend on active sheet/selection in a way users will find surprising?


## NativeAOT interop implications

NativeAOT makes the Excel-DNA add-in more like a native binary from Excel's point of view, but it does not remove Excel's C API rules. Functions, commands, macro contexts, thread-safety flags, bitness, and registration still matter.

Additional NativeAOT-specific points:

- The current Excel-DNA NativeAOT preview produces native 64-bit XLLs, so 64-bit Excel is the baseline.
- NativeAOT has restrictions around ordinary Windows COM support. For Excel object model access in NativeAOT preview, prefer the Excel-DNA `DynamicApplication`/`IDynamic` abstraction shown in the NativeAOT docs.
- P/Invoke can work under NativeAOT, but direct P/Invoke and static native linking change startup binding and dependency search behavior. Test native dependencies on clean machines.
- Avoid dynamic assembly loading and runtime-generated interop shims.
- C API and COM calls still must respect Excel context rules; NativeAOT does not make workbook mutation safe from UDF calculation or background threads.

When porting C API-adjacent code to NativeAOT, preserve the Excel-DNA safety rules first, then handle AOT warnings and native dependency packaging.

## Separate future C API skill

A dedicated C API skill should later cover:

- `XLOPER` / `XLOPER12` memory rules.
- `Excel4`, `Excel12`, and callbacks.
- Type text strings for `xlfRegister`.
- Multi-threaded recalculation flags.
- Async handles and RTD internals.
- `xlAutoOpen`/`xlAutoClose` lifecycle.
- Command vs function registration.
- Cluster-safe and macro-sheet behaviors.

For now, this Excel-DNA skill includes the C API knowledge needed to build correct managed XLL add-ins.
