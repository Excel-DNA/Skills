using ExcelDna.Integration;

namespace MyCompany.MyAddIn;

public static class Functions
{
    [ExcelFunction(Name = "MY.RUNTIME", Description = "Returns the runtime description.", Category = "My Add-in")]
    public static string Runtime()
    {
        return System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription;
    }
}
