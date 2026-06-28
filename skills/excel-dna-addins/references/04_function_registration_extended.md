# Function Registration and Extended Patterns

## Contents

- [Default registration](#default-registration)
- [Attribute-based metadata](#attribute-based-metadata)
- [Optional/default parameters](#optionaldefault-parameters)
- [Parameter conversion principles](#parameter-conversion-principles)
- [Function execution handlers](#function-execution-handlers)
- [Registration transforms](#registration-transforms)
- [Object handles](#object-handles)
- [Naming and namespace transforms](#naming-and-namespace-transforms)
- [Caching](#caching)
- [Error wrapping across a library](#error-wrapping-across-a-library)
- [NativeAOT registration constraints](#nativeaot-registration-constraints)

## Default registration

Excel-DNA registers public static functions as Excel UDFs by default, depending on project settings. For production add-ins, prefer explicit exports:

```xml
<PropertyGroup>
  <ExcelAddInExplicitExports>true</ExcelAddInExplicitExports>
</PropertyGroup>
```

Then only functions marked with `[ExcelFunction]` are exposed.

Why this matters:

- Prevents accidental public helper methods from becoming worksheet functions.
- Makes the function surface auditable.
- Keeps metadata complete.

## Attribute-based metadata

```csharp
[ExcelFunction(
    Name = "MYCO.FUNC",
    Description = "Does something useful.",
    Category = "MyCo",
    HelpTopic = "https://example.invalid/help/myco-func",
    IsThreadSafe = true)]
public static double Func(
    [ExcelArgument(Name = "x", Description = "Input value")] double x)
{
    return x;
}
```

Review attributes during code generation:

- `Name` should be stable and collision-resistant.
- `Description` should be user-facing and concise.
- `Category` should group functions consistently.
- `IsThreadSafe` should be true only for safe functions.
- `IsVolatile` should be rare and justified.
- `IsMacroType` should be rare and justified.

`HelpTopic` note: Excel's `xlfRegister` requires `path!HelpContextID` (e.g., `https://example.com/help!0`). On the standard registration path Excel-DNA's loader appends `!0` automatically when the `HelpTopic` value starts with `http://`, `https://`, or `file://` and is missing the suffix, so the bare URL above works for ordinary `[ExcelFunction]` registration. **On the Registration pipeline path** (functions wrapped through `ExcelRegistration.RegisterFunctions(...)` — the same pipeline used to install execution handlers and registration transforms), that auto-fix is bypassed, and you must include `!0` in the attribute value yourself. If you mix `HelpTopic` URLs with handlers, write `"https://example.com/help!0"` explicitly.

## Optional/default parameters

Modern Excel-DNA registration supports more optional/default parameter patterns than older examples. Still be explicit about how omitted and empty values behave.

Recommended style for stable public functions:

```csharp
[ExcelFunction(Name = "MYCO.POWER", Description = "Raises x to p; p defaults to 2.")]
public static double Power(double x, double p = 2.0) => Math.Pow(x, p);
```

When compatibility or nuanced missing/empty behavior matters:

```csharp
[ExcelFunction(Name = "MYCO.POWER2", Description = "Raises x to p; p defaults to 2.")]
public static object Power2(double x, object p)
{
    double exponent = p is ExcelMissing or ExcelEmpty ? 2.0 : Convert.ToDouble(p);
    return Math.Pow(x, exponent);
}
```

## Parameter conversion principles

- Keep public functions typed where possible.
- Use `object` only when the Excel value shape really varies.
- Use `object[,]` for ranges.
- Use `ExcelReference` only when range identity or C API operations are needed.
- Treat `DateTime` carefully; Excel serial dates and locale display can confuse users.
- Avoid accepting arbitrary strings when a typed number/date/bool would be clearer.

## Function execution handlers

Excel-DNA's newer registration model includes function execution handler extension points. These enable aspect-like behavior across many functions:

- logging
- timing
- error mapping
- caching
- validation
- telemetry
- authorization / feature flags
- dependency injection of services

The handler types live in the `ExcelDna.Registration` namespace (shipped inside the `ExcelDna.Integration` assembly in current Excel-DNA versions).

```csharp
using ExcelDna.Integration;
using ExcelDna.Registration;
using System.Diagnostics;

public sealed class TimingHandler : FunctionExecutionHandler
{
    public override void OnEntry(FunctionExecutionArgs args)
    {
        args.Tag = Stopwatch.StartNew();
    }

    public override void OnSuccess(FunctionExecutionArgs args)
    {
        if (args.Tag is Stopwatch sw)
        {
            sw.Stop();
            // Log args.FunctionName and sw.Elapsed.
        }
    }

    public override void OnException(FunctionExecutionArgs args)
    {
        // Map known exceptions to ExcelError, log details elsewhere.
        args.ReturnValue = ExcelError.ExcelErrorValue;
        args.FlowBehavior = FlowBehavior.Return;
    }
}
```

Key points to keep this snippet correct against the current source:

- `FunctionExecutionArgs` has no `Handled` property. Flow control is via `args.FlowBehavior`, with `FlowBehavior.Return` (use `args.ReturnValue`), `FlowBehavior.RethrowException`, `FlowBehavior.ThrowException`, and `FlowBehavior.Default`.
- Available `FunctionExecutionArgs` members: `FunctionName`, `Arguments` (read-only), `ReturnValue`, `Exception`, `FlowBehavior`, `Tag`.
- Virtuals are `OnEntry`, `OnSuccess`, `OnException`, `OnExit`. The class also implements `IFunctionExecutionHandler` if you prefer the interface form.

Apply cross-cutting behavior at registration/execution rather than copy/pasting wrappers into every UDF.

## Registration transforms

Extended registration can transform functions before Excel registration. This is useful for:

- adding generated metadata
- wrapping functions in handlers
- converting custom parameter types
- registering async functions
- registering object handles
- applying consistent prefixes/categories
- hiding internal functions

Design rule: keep transformations predictable and inspectable. If a function's Excel name, category, behavior, or threading changes through registration code, document that transformation in one place.

## Object handles

Object handles let formulas work with complex .NET objects without serializing the whole object into worksheet cells.

Use handles when:

- A function creates an expensive object such as a model, curve, connection, parsed document, or simulation state.
- Other functions need to consume that object repeatedly.
- Returning a scalar or 2D array would lose identity or be too expensive.
- Updates should propagate through RTD/observable behavior.

Example conceptual surface:

```excel
=MYCO.CURVE.CREATE(dates, rates)       -> "Curve:42"
=MYCO.CURVE.DISCOUNT("Curve:42", date) -> 0.9731
=MYCO.CURVE.INFO("Curve:42")           -> spill/table of metadata
```

Guidelines:

- Make handle strings recognizable and stable enough for users.
- Include a way to inspect handle type/status.
- Define disposal/lifecycle semantics.
- Avoid memory leaks from abandoned handles.
- Avoid hiding calculation dependencies: if a handle depends on inputs, update it when inputs change.
- Consider versioned handles so dependent formulas recalculate when the underlying object changes.

## Naming and namespace transforms

For larger libraries, consider a registration pass that enforces:

- Prefix: `MYCO.`
- Category based on namespace/class.
- Default descriptions from XML docs or source annotations.
- Consistent error handling.

But avoid surprising names generated from implementation details that may change.

## Caching

Caching in UDFs is a double-edged sword. Safe patterns:

- Cache immutable results keyed by all inputs.
- Clear caches on workbook close/reopen or add-in reload when appropriate.
- Expose diagnostics for cache size/hits.
- Avoid caching values dependent on hidden global state unless the function is volatile or has explicit dependency inputs.

Unsafe patterns:

- Cache by function name only.
- Cache mutable objects without thread safety.
- Cache external service results indefinitely without user-visible refresh controls.
- Cache workbook/cell-specific data without considering workbook identity.

## Error wrapping across a library

Use a function execution handler to centralize:

- exception-to-Excel-error mapping
- diagnostic logging
- optional debug mode returning more verbose information
- user-friendly messages for common bad inputs

This is better than letting arbitrary exceptions leak into Excel cells.


## NativeAOT registration constraints

NativeAOT changes the implementation constraints around registration and wrapping. The worksheet surface can still look like ordinary Excel-DNA functions, but the implementation must be AOT-compatible.

Guidelines for agents:

- Prefer explicit static functions with `[ExcelFunction]` over convention-only discovery.
- Keep registration transforms deterministic and based on statically known functions/types.
- Avoid transform code that depends on dynamic assembly loading, `Reflection.Emit`, expression compilation, or runtime-generated proxies.
- Function execution handlers are shown in NativeAOT preview examples, but handler code and any logging/telemetry dependencies must be AOT-compatible.
- Object handles are shown in NativeAOT preview examples. Keep handle types statically known and avoid reflection-heavy factories.
- User-defined conversions are shown in preview examples. Keep conversion methods simple and attributed, not discovered through broad reflection scans.

If a NativeAOT add-in fails with missing native code or metadata for a large generic expression/delegate shape, suspect registration wrapping, async wrapper generation, or custom conversion code first. Reduce to a minimal UDF, then add wrappers/conversions back one feature at a time.
