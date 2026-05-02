# Async/Streaming Template Notes

Use `Task<T>` for one-shot asynchronous functions. Use `IObservable<T>` or explicit RTD utilities for repeated updates. Include timeouts and cancellation for external operations.
