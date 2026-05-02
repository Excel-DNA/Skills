# Async, Streaming, RTD, and Batching

## Contents

- [When async is useful](#when-async-is-useful)
- [Task-based async functions](#task-based-async-functions)
- [Cancellation token pattern](#cancellation-token-pattern)
- [Streaming functions with `IObservable<T>`](#streaming-functions-with-iobservablet)
- [`ExcelAsyncUtil.Run`](#excelasyncutilrun)
- [RTD-style async vs native Excel async](#rtd-style-async-vs-native-excel-async)
- [Batching](#batching)
- [Lossless vs latest-value streams](#lossless-vs-latest-value-streams)
- [Error handling in async functions](#error-handling-in-async-functions)
- [Threading and UI](#threading-and-ui)
- [NativeAOT async cautions](#nativeaot-async-cautions)

## When async is useful

Use async for I/O-bound or long-running operations that should not block Excel calculation unnecessarily:

- HTTP/API calls
- database queries
- slow file/network operations
- service calls
- computations offloaded to a worker process

Do not use async as a performance bandage for pure CPU-bound functions. For CPU-bound local work, prefer efficient code, arrays, caching, and thread-safe recalculation.

## Task-based async functions

Modern Excel-DNA registration can expose `Task<T>` returning functions. This is the preferred simple pattern.

```csharp
using ExcelDna.Integration;
using System.Net.Http;

public static class AsyncFunctions
{
    private static readonly HttpClient Http = new();

    [ExcelFunction(Name = "MYCO.HTTP.GET", Description = "Gets text from a URL asynchronously.")]
    public static async Task<string> HttpGet(
        [ExcelArgument(Name = "url", Description = "URL to fetch")] string url)
    {
        return await Http.GetStringAsync(url).ConfigureAwait(false);
    }
}
```

Design notes:

- Include all input values in the function signature so Excel-DNA can identify calls correctly.
- Return stable, Excel-friendly results.
- Map errors to user-visible status/error values where possible.
- Use `CancellationToken` when supported by the registration pattern and the operation can be cancelled.

## Cancellation token pattern

For async calls that can be cancelled when the formula is deleted or recalculation changes:

```csharp
[ExcelFunction(Name = "MYCO.WAIT", Description = "Waits asynchronously, then returns a message.")]
public static async Task<string> WaitAsync(int milliseconds, CancellationToken cancellationToken)
{
    await Task.Delay(milliseconds, cancellationToken).ConfigureAwait(false);
    return "done";
}
```

The cancellation token should be last and should not be exposed as a worksheet argument.

## Streaming functions with `IObservable<T>`

Streaming functions are appropriate for values that update over time:

- market data
- clocks/timers
- progress/status feeds
- external subscriptions
- model state updates

Conceptual example:

```csharp
[ExcelFunction(Name = "MYCO.CLOCK", Description = "Streams the current time.")]
public static IObservable<string> Clock()
{
    return new ClockObservable();
}
```

Streaming functions are typically backed by RTD mechanics. The observable should:

- publish initial value quickly
- stop when disposed/unsubscribed
- avoid unbounded queues
- avoid updating too frequently for Excel to remain responsive
- report errors deliberately

## `ExcelAsyncUtil.Run`

Older and still useful pattern:

```csharp
[ExcelFunction(Name = "MYCO.SLOWADD", Description = "Adds after a delay asynchronously.")]
public static object SlowAdd(double x, double y)
{
    return ExcelAsyncUtil.Run(
        "MYCO.SLOWADD",
        new object[] { x, y },
        () =>
        {
            Thread.Sleep(1000);
            return x + y;
        });
}
```

Rules:

- The function identity plus parameters must uniquely identify the async call.
- Include every input argument in the identity arguments.
- Do not perform workbook mutation in the async delegate.
- Prefer `Task<T>` for new code unless the explicit RTD pattern is required.

## RTD-style async vs native Excel async

Excel-DNA async historically used RTD-style mechanics. Advantages include compatibility and streaming flexibility. Task-based async gives a simpler surface while using appropriate underlying mechanisms.

Use explicit RTD/observable patterns when:

- The value changes repeatedly.
- You need custom subscription lifecycle control.
- You are modeling a real-time feed.

Use task-based async when:

- The calculation eventually returns one result.
- The operation maps naturally to `Task<T>`.
- You want the simplest public UDF implementation.

## Batching

Batching reduces duplicate external calls when many cells call related functions.

Patterns:

- Use an in-memory request queue keyed by request arguments.
- Coalesce requests during a short time window.
- Make one external call and distribute results.
- Return intermediate status while the batch is in flight.

Hazards:

- Batching can make dependencies less obvious.
- Shared queues must be thread-safe.
- User-visible latency may increase if the batch window is too long.
- Cancellation and workbook recalculation can leave stale requests.

## Lossless vs latest-value streams

For high-frequency streams, Excel often needs the latest value, not every tick. Lossless streams may be necessary for audit or event processing, but they can overwhelm Excel if the source updates faster than Excel can recalculate.

Use latest-value streams for market quotes/status. Use lossless streams only when every update is semantically required and bounded.

## Error handling in async functions

Avoid letting raw exceptions appear unpredictably. A robust async function returns:

- a final scalar/array result
- an Excel error for ordinary invalid input
- a clear status string or error code for external service failures, when that is the product design

Centralize logging outside Excel cells.

## Threading and UI

Async work should not call Excel COM APIs on background threads. If a completion must update Excel state beyond returning the function value, queue a macro-context action and keep it separate from the UDF result path.


## NativeAOT async cautions

Excel-DNA NativeAOT preview examples include async functions and task-returning functions, so async is not automatically excluded. The risk is usually not `Task<T>` itself, but the wrappers, dependencies, and dynamic patterns around it.

For NativeAOT preview async functions:

- Prefer simple `Task<T>` signatures first.
- Add cancellation only when supported and tested by the current preview path.
- Avoid expression-tree-based or runtime-generated async wrappers.
- Treat missing native code/metadata errors as AOT compatibility failures, not ordinary Excel-DNA load failures.
- Test formula deletion, recalculation, workbook close, and add-in unload.
- Keep workbook mutation out of the background task.

For CPU-bound work, NativeAOT is not a substitute for algorithm design. Excel multi-threaded recalculation, batching, array formulas, and object handles may matter more than AOT compilation. Benchmark repeated recalculation separately from cold startup.
