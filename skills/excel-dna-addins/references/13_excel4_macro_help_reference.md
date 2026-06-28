# Excel 4 Macro Help Reference

## Contents

- [When to use this reference](#when-to-use-this-reference)
- [Local source material](#local-source-material)
- [Excel-DNA relevance](#excel-dna-relevance)
- [Execution context rules](#execution-context-rules)
- [High-value XLM topics](#high-value-xlm-topics)
- [Using XLM from Excel-DNA](#using-xlm-from-excel-dna)
- [Safety and modernization guidance](#safety-and-modernization-guidance)

## When to use this reference

Open this file when a user asks about Excel 4 macros, XLM macro sheets, macro-sheet functions, `GET.CELL`, `GET.WORKBOOK`, `CALLER`, `EVALUATE`, `REGISTER`, `CALL`, macro commands, `ExecuteExcel4Macro`, or why an `XlCall` call is allowed in one Excel-DNA context but fails in another.

Do not treat XLM as the default path for ordinary Excel-DNA development. Prefer normal UDFs, ribbon callbacks, commands, COM object model calls in valid contexts, and Excel-DNA abstractions. Reach for XLM when the user is maintaining legacy XLM behavior, interrogating Excel state exposed through macro-sheet functions, or working near the Excel C API/XLL boundary.

## Local source material

The local reference corpus is under `C:/Work/ExcelDna/MacroHelp`.

Important files:

- `TerryAney-copilot.excel.macro.reference/copilot.excel.macro.reference.md`: text-friendly Markdown conversion of the Excel 4.0 Macro Reference.
- `TerryAney-copilot.excel.macro.reference/Excel 4 Macro Reference.pdf`: PDF source paired with the Markdown conversion.
- `Excel 4 Macro Reference.docx` and `Excel 4 Macros.docx`: Word copies.
- `XLMACR8.chm`, `XLMACR8.chw`, and `XLMACR8.zip`: compiled help/source archive copies.
- `Readme.txt`: points to Theo Heselmans' archived macro help post.

Use the Markdown file first for searching. Use the PDF/CHM only when formatting, tables, or images in the conversion are suspect.

## Excel-DNA relevance

Excel-DNA exposes the same Excel calculation and command engine that XLM macros used, but through safer managed surfaces:

- `XlCall.Excel` and `XlCall.TryExcel` call Excel C API functions and commands by numeric/function constants.
- `ExcelReference` carries a real Excel reference when a macro-sheet function needs reference identity rather than only values.
- `[ExcelFunction(IsMacroType = true)]` can be required for UDFs that need macro-sheet information functions, but it should be used only for a specific reason.
- `ExcelAsyncUtil.QueueAsMacro` schedules work in a macro-capable context for state-changing operations.
- `ExcelDnaUtil.Application` gives the Excel COM object model in valid UI/macro contexts; the COM object model also exposes `ExecuteExcel4Macro` for string-based XLM evaluation when that is the least bad interop path.

## Execution context rules

Excel's XLM function set includes both worksheet-like information functions and command macros that mutate Excel state. Context matters more than syntax.

- Normal worksheet UDFs should remain deterministic and side-effect free.
- Thread-safe UDFs must not call Excel COM and should not use macro commands or active-selection-dependent XLM calls.
- Macro-sheet information functions such as `GET.CELL`, `GET.WORKBOOK`, and `CALLER` may require macro-type registration or may fail in ordinary UDF contexts.
- Command macros such as `FORMULA`, `SELECT`, `ON.TIME`, and `RUN` belong in command, ribbon, or queued macro flows, not in calculation.
- Prefer `TryExcel` when probing Excel state from uncertain contexts; return an intentional Excel error or fallback rather than letting a failed C API call escape as an exception.

## High-value XLM topics

| Topic | XLM entries | Why it matters for Excel-DNA |
|---|---|---|
| Caller identity | `CALLER` | Explains the macro-sheet model behind `xlfCaller` and why caller shape differs for cells, menu commands, toolbar commands, and objects. |
| Cell metadata | `GET.CELL`, `GET.FORMULA` | Useful for legacy add-ins that need formatting, formula text, protection, or other cell metadata not passed as normal UDF values. |
| Workbook/window/document state | `GET.WORKBOOK`, `GET.WINDOW`, `GET.DOCUMENT`, `GET.WORKSPACE` | Maps to common diagnostic and compatibility probes, but many results are active-window dependent. |
| Name and formula evaluation | `GET.NAME`, `EVALUATE`, `FORMULA`, `FORMULA.FILL` | Helps when maintaining legacy defined-name tricks or evaluating text formulas. Prefer native Excel formulas or managed parsing when possible. |
| Function registration | `REGISTER`, `REGISTER.ID`, `UNREGISTER`, `ARGUMENT`, `VOLATILE` | Historical basis for XLL function metadata, type text, macro type flags, volatility, and Function Wizard help. Excel-DNA normally handles this through attributes and registration code. |
| Macro control flow | `RUN`, `RETURN`, `HALT`, `PAUSE`, `ON.TIME`, `ON.KEY`, `ON.WINDOW`, `ON.SHEET`, `ON.RECALC` | Important for migrating old XLM automation. For new Excel-DNA code, prefer explicit commands, events, RTD/async patterns, or queued macros. |
| UI and selection commands | `ALERT`, `MESSAGE`, `SELECT`, `ACTIVATE`, `SET.VALUE`, `WORKBOOK.*`, `WINDOW.*` | These are side-effecting macro commands. Use ribbon callbacks, commands, queued macros, or COM automation instead for new code. |
| External native calls | `CALL`, `REGISTER` | Legacy unmanaged interop path. Avoid recommending this for generated Excel-DNA projects; use managed code, vetted P/Invoke, or a native XLL intentionally designed for the task. |

## Using XLM from Excel-DNA

Use `XlCall.TryExcel` for C API calls that may be context-sensitive:

```csharp
[ExcelFunction(Name = "MYCO.CALLER.TEXT", IsMacroType = true)]
public static object CallerText()
{
    var status = XlCall.TryExcel(XlCall.xlfCaller, out var caller);
    return status == XlCall.XlReturn.XlReturnSuccess
        ? caller?.ToString() ?? ExcelEmpty.Value
        : ExcelError.ExcelErrorValue;
}
```

Use `QueueAsMacro` for workbook mutation:

```csharp
public static void SelectA1Later()
{
    ExcelAsyncUtil.QueueAsMacro(() =>
    {
        XlCall.Excel(XlCall.xlcSelect, "R1C1");
    });
}
```

Use COM `ExecuteExcel4Macro` only from a valid macro/ribbon/command context when a specific XLM string expression is required:

```csharp
public static object EvaluateXlm(string expression)
{
    dynamic excel = ExcelDnaUtil.Application;
    return excel.ExecuteExcel4Macro(expression);
}
```

Do not pass untrusted user input to `ExecuteExcel4Macro`, `EVALUATE`, `RUN`, `CALL`, or `REGISTER`. Treat them as code execution surfaces.

## Safety and modernization guidance

- Prefer `[ExcelFunction(IsVolatile = true)]` over manually invoking XLM `VOLATILE` in new Excel-DNA UDFs.
- Prefer `[ExcelFunction]` and `[ExcelArgument]` metadata over XLM `REGISTER`/`ARGUMENT` for Function Wizard content.
- Prefer normal Excel-DNA registration and packing over hand-authored `REGISTER` calls.
- Avoid active-selection-dependent XLM calls in UDFs. If a function depends on active sheet/window state, document the volatility and calculation implications.
- Verify legacy XLM functions in the target Excel version. The MacroHelp corpus includes old menu, toolbar, mail, and dialog-sheet features that may be obsolete or behave differently in current Microsoft 365 Excel.
- For NativeAOT preview projects, be extra conservative: direct C API calls still obey Excel context rules, and COM/XLM bridges must be tested under the exact Excel bitness and runtime model.
