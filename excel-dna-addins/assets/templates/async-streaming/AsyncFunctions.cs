using ExcelDna.Integration;
using System;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;

namespace MyCompany.MyAddIn;

public static class AsyncFunctions
{
    private static readonly HttpClient Http = new() { Timeout = TimeSpan.FromSeconds(15) };

    [ExcelFunction(Name = "MY.HTTP.GET", Description = "Gets text from a URL asynchronously.", Category = "My Add-in")]
    public static async Task<object> HttpGet(
        [ExcelArgument(Name = "url", Description = "URL to fetch")] string url,
        CancellationToken cancellationToken)
    {
        try
        {
            using var response = await Http.GetAsync(url, cancellationToken).ConfigureAwait(false);
            response.EnsureSuccessStatusCode();
            return await response.Content.ReadAsStringAsync().ConfigureAwait(false);
        }
        catch (OperationCanceledException)
        {
            return ExcelError.ExcelErrorNA;
        }
        catch
        {
            return ExcelError.ExcelErrorValue;
        }
    }
}
