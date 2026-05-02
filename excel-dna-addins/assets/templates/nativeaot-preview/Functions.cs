using ExcelDna.Integration;

namespace MyNativeAddin;

public static class Functions
{
    [ExcelFunction(
        Name = "NATIVE.ADD",
        Description = "Adds two numbers from a NativeAOT Excel-DNA add-in.",
        Category = "NativeAOT Preview",
        IsThreadSafe = true)]
    public static double Add(
        [ExcelArgument(Name = "x", Description = "First number")] double x,
        [ExcelArgument(Name = "y", Description = "Second number")] double y)
    {
        return x + y;
    }

    [ExcelFunction(
        Name = "NATIVE.HELLO",
        Description = "Returns a greeting from a NativeAOT Excel-DNA add-in.",
        Category = "NativeAOT Preview")]
    public static string Hello(
        [ExcelArgument(Name = "name", Description = "Name to greet")] string name)
    {
        return $"Hello from NativeAOT, {name}!";
    }
}
