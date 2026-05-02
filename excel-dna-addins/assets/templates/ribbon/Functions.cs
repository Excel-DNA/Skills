using ExcelDna.Integration;

namespace MyCompany.MyAddIn;

public static class Functions
{
    [ExcelFunction(Name = "MY.ADD", Description = "Adds two numbers.", Category = "My Add-in", IsThreadSafe = true)]
    public static double Add(double x, double y) => x + y;
}
