using ExcelDna.Integration;

namespace MyCompany.MyAddIn;

public static class Functions
{
    [ExcelFunction(Name = "MY.ADD", Description = "Adds two numbers.", Category = "My Add-in", IsThreadSafe = true)]
    public static double Add(
        [ExcelArgument(Name = "x", Description = "First number")] double x,
        [ExcelArgument(Name = "y", Description = "Second number")] double y)
    {
        return x + y;
    }
}
