# UDF Core Patterns

## Contents

- [Why UDFs are the core Excel-DNA mechanism](#why-udfs-are-the-core-excel-dna-mechanism)
- [Function shape](#function-shape)
- [Supported parameter and return concepts](#supported-parameter-and-return-concepts)
- [Range and array handling](#range-and-array-handling)
- [Excel special values](#excel-special-values)
- [Thread safety](#thread-safety)
- [Volatile functions](#volatile-functions)
- [Macro-type functions](#macro-type-functions)
- [Side effects and workbook mutation](#side-effects-and-workbook-mutation)
- [Function metadata](#function-metadata)
- [Error strategy](#error-strategy)
- [Performance patterns](#performance-patterns)
- [NativeAOT-specific UDF guidance](#nativeaot-specific-udf-guidance)

## Why UDFs are the core Excel-DNA mechanism

A worksheet UDF lets users stay inside Excel's primary modeling language: formulas. This is powerful because the function participates in recalculation, dependency graphs, workbook auditing, copy/fill operations, table formulas, array formulas, and ordinary spreadsheet collaboration.

Compared with buttons or macros, UDFs are composable. A single function can be used thousands of times in a workbook model, combined with native Excel functions, and recalculated by Excel when inputs change. This is why Excel-DNA should usually be designed UDF-first, with ribbons and task panes supporting function discovery, configuration, and workflow rather than replacing formulas.

## Function shape

A typical production function:

```csharp
[ExcelFunction(
    Name = "MYCO.DISCOUNT",
    Description = "Calculates the discounted value of a cash flow.",
    Category = "MyCo Finance",
    IsThreadSafe = true)]
public static object Discount(
    [ExcelArgument(Name = "amount", Description = "Cash flow amount")] double amount,
    [ExcelArgument(Name = "rate", Description = "Annual discount rate")] double rate,
    [ExcelArgument(Name = "years", Description = "Number of years")] double years)
{
    if (years < 0) return ExcelError.ExcelErrorNum;
    return amount / Math.Pow(1.0 + rate, years);
}
```

For functions returning error values, prefer return type `object` if the function sometimes returns an Excel error and sometimes returns a scalar:

```csharp
[ExcelFunction(Name = "MYCO.SAFELOG", Description = "Natural log with Excel error for invalid input.")]
public static object SafeLog(double x)
{
    return x > 0 ? Math.Log(x) : ExcelError.ExcelErrorNum;
}
```

## Supported parameter and return concepts

Common practical types:

- `double`, `int`, `bool`, `string`
- `DateTime` where appropriate, with clear conversion semantics
- `object` for mixed Excel values
- `object[,]` for ranges and arrays
- `ExcelReference` for direct references when the function needs the range identity, not just values
- `ExcelError` for Excel error return values
- `Task<T>` and `IObservable<T>` for async/streaming behavior

## Range and array handling

For rectangular ranges, use `object[,]`:

```csharp
[ExcelFunction(Name = "MYCO.SUMTEXTNUMBERS", Description = "Sums numeric cells in a range.")]
public static double SumTextNumbers(object[,] values)
{
    double total = 0;
    foreach (var value in values)
    {
        switch (value)
        {
            case double d:
                total += d;
                break;
            case string s when double.TryParse(s, out var parsed):
                total += parsed;
                break;
        }
    }
    return total;
}
```

Rules:

- Treat input arrays as 2D even for a single row or column.
- Preserve dimensions when returning arrays.
- Return `ExcelError.ExcelErrorNA` or `ExcelEmpty.Value` deliberately for missing elements.
- Document orientation: row vector vs column vector.

## Excel special values

Excel-DNA exposes special values such as:

- `ExcelMissing.Value` for omitted optional arguments.
- `ExcelEmpty.Value` for empty cells.
- `ExcelError.ExcelErrorNA`, `ExcelError.ExcelErrorValue`, `ExcelError.ExcelErrorNum`, and related error values.

A robust optional argument pattern:

```csharp
[ExcelFunction(Name = "MYCO.REPEAT", Description = "Repeats text.")]
public static object Repeat(
    string text,
    [ExcelArgument(Description = "Repeat count; defaults to 2.")] object count)
{
    int n = count is ExcelMissing or ExcelEmpty ? 2 : Convert.ToInt32(count);
    if (n < 0) return ExcelError.ExcelErrorNum;
    return string.Concat(Enumerable.Repeat(text, n));
}
```

In newer registration paths, optional/default parameters can be supported more directly. Still understand `ExcelMissing` because it appears in legacy projects and explicit conversion paths.

## Thread safety

Mark `IsThreadSafe=true` only when the function is genuinely safe under Excel multi-threaded recalculation:

- No mutable shared state without synchronization.
- No Excel COM object model access.
- No UI access.
- No dependence on calculation order except through worksheet dependencies.
- No static caches that mutate unsafely.

Good thread-safe UDFs are pure or use immutable/cache-safe state.

## Volatile functions

`IsVolatile=true` makes Excel recalculate the function more often. Use sparingly. Volatility can destroy workbook performance.

Use volatility for functions genuinely dependent on non-formula state, such as time, environment, or external cache status. Prefer explicit inputs that make dependencies visible.

## Macro-type functions

`IsMacroType=true` gives access to a broader set of macro-sheet functionality and can affect recalculation behavior. Use only when needed, and document the reason. It is not a general solution for side effects in worksheet formulas.

## Side effects and workbook mutation

Normal UDFs should not mutate the workbook, show UI, write cells, save files, or use arbitrary COM automation. Excel may call UDFs during calculation in contexts where side effects are unsafe or surprising.

For side effects:

- Use ribbon callbacks, commands, or macros.
- Use `ExcelAsyncUtil.QueueAsMacro` to transition work to a macro context.
- Separate calculation from action: UDF returns data/status; command performs mutation.

## Function metadata

Always provide:

- `Name`
- `Description`
- `Category`
- `ExcelArgument` names and descriptions

Metadata improves Function Wizard behavior, IntelliSense, examples, and maintainability.

## Error strategy

Use Excel errors rather than throwing exceptions for ordinary user-facing validation failures. Reserve thrown exceptions for programming errors and catch/map them through a function execution handler where possible.

A consistent map:

| Condition | Suggested return |
|---|---|
| Invalid argument | `#VALUE!` or `#NUM!` |
| Missing lookup | `#N/A` |
| Unsupported range shape | `#VALUE!` |
| External service temporarily unavailable | `#N/A` or clear status string, depending domain |
| Internal bug | custom error wrapper or `#VALUE!` plus logging |

## Performance patterns

- Avoid per-cell expensive setup; cache immutable resources.
- Avoid accessing Excel COM from UDFs.
- Batch external calls when possible.
- Use arrays to reduce formula count when appropriate.
- Use async for I/O-bound work, not CPU-bound local math.
- Use object handles for large objects used across multiple functions.
- Make thread-safe pure functions where they benefit from Excel multi-threaded recalc.


## NativeAOT-specific UDF guidance

For NativeAOT preview projects, keep the Excel-facing function surface especially simple and statically discoverable:

- Use explicit `[ExcelFunction]` attributes and complete `[ExcelArgument]` metadata.
- Prefer primitive/scalar, `object`, `object[,]`, and other Excel-DNA-supported types shown in the NativeAOT preview docs.
- Avoid dynamic registration machinery that builds wrappers through runtime code generation or unbounded reflection.
- Avoid function implementations that depend on dynamically loaded assemblies or plugins.
- Publish often and treat AOT/trimming warnings as defects.
- Benchmark in Excel before claiming NativeAOT is faster. Cold load, first calculation, repeated calculation, memory use, and binary size are separate metrics.

NativeAOT is promising for UDF-first add-ins because the worksheet-facing surface is often static and closed-world friendly. It is less promising for add-ins whose value is rich dynamic UI, plugin loading, or runtime code-generation extensibility.
