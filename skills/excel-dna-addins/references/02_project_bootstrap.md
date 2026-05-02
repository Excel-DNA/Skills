# Project Bootstrap

## Contents

- [Minimal modern C# project](#minimal-modern-c-project)
- [Minimal function file](#minimal-function-file)
- [NativeAOT preview project](#nativeaot-preview-project)
- [Debugging and running](#debugging-and-running)
- [Project file properties worth knowing](#project-file-properties-worth-knowing)
- [`.dna` file role](#dna-file-role)
- [Naming strategy](#naming-strategy)
- [Architecture checklist](#architecture-checklist)
- [Use the scaffold script](#use-the-scaffold-script)

## Minimal modern C# project

A minimal production-shaped Excel-DNA add-in is an SDK-style class library with an Excel-DNA package reference, a Windows-compatible target framework, and explicitly exported UDFs.

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFramework>net48</TargetFramework>
    <LangVersion>latest</LangVersion>
    <ExcelAddInExplicitExports>true</ExcelAddInExplicitExports>
    <RunExcelDnaPack>true</RunExcelDnaPack>
  </PropertyGroup>

  <ItemGroup>
    <PackageReference Include="ExcelDna.AddIn" Version="1.9.0" />
  </ItemGroup>
</Project>
```

For modern .NET on controlled desktops:

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFramework>net8.0-windows</TargetFramework>
    <UseWindowsForms>true</UseWindowsForms>
    <Nullable>enable</Nullable>
    <LangVersion>latest</LangVersion>
    <ExcelAddInExplicitExports>true</ExcelAddInExplicitExports>
    <RunExcelDnaPack>true</RunExcelDnaPack>
    <RollForward>LatestMajor</RollForward>
  </PropertyGroup>

  <ItemGroup>
    <PackageReference Include="ExcelDna.AddIn" Version="1.9.0" />
  </ItemGroup>
</Project>
```

Multi-targeting is useful for libraries that want broad stable distribution and modern development:

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFrameworks>net48;net8.0-windows</TargetFrameworks>
    <Nullable>enable</Nullable>
    <LangVersion>latest</LangVersion>
    <ExcelAddInExplicitExports>true</ExcelAddInExplicitExports>
    <RunExcelDnaPack>true</RunExcelDnaPack>
  </PropertyGroup>

  <ItemGroup>
    <PackageReference Include="ExcelDna.AddIn" Version="1.9.0" />
  </ItemGroup>
</Project>
```

## Minimal function file

```csharp
using ExcelDna.Integration;

namespace MyCompany.MyAddIn;

public static class Functions
{
    [ExcelFunction(
        Name = "MY.ADD",
        Description = "Adds two numbers.",
        Category = "My Add-in")]
    public static double Add(
        [ExcelArgument(Name = "x", Description = "First number")] double x,
        [ExcelArgument(Name = "y", Description = "Second number")] double y)
    {
        return x + y;
    }
}
```


## NativeAOT preview project

NativeAOT is not a normal managed add-in target. Use a separate project or conditional project sections, and read `references/11_native_aot_preview.md` before recommending it.

Minimal NativeAOT preview project shape:

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFramework>net10.0-windows</TargetFramework>
    <RuntimeIdentifier>win-x64</RuntimeIdentifier>
    <PublishAot>true</PublishAot>
    <Nullable>enable</Nullable>
    <ImplicitUsings>enable</ImplicitUsings>
    <LangVersion>latest</LangVersion>
    <ExcelAddInExplicitExports>true</ExcelAddInExplicitExports>
  </PropertyGroup>

  <ItemGroup>
    <PackageReference Include="ExcelDna.AddIn.NativeAOT" Version="1.10.0-preview4" />
  </ItemGroup>
</Project>
```

Build/publish:

```bash
dotnet restore -r win-x64
dotnet publish -c Release -r win-x64
```

Load the published `*-AddIn64.xll` from the RID-specific publish folder in 64-bit Excel. Do not load the ordinary build DLL or assume the normal managed output path.

NativeAOT bootstrap rules:

- Prefer `ExcelDna.AddIn.NativeAOT` alone in the NativeAOT project.
- Avoid unconditionally referencing both `ExcelDna.AddIn` and `ExcelDna.AddIn.NativeAOT`.
- Keep all Excel-DNA preview package versions aligned if conditional multi-targeting is unavoidable.
- Do not add `ExcelDna.IntelliSense` to NativeAOT preview projects unless current source/tests prove support.
- Treat all AOT/trimming warnings as defects until investigated.

## Debugging and running

Typical flow:

1. Build the project.
2. Use the generated `.xll` from the build/publish output.
3. During development, launch Excel with the add-in loaded, or configure Visual Studio debugging through Excel-DNA build/debug properties.
4. Confirm the function in a workbook formula: `=MY.ADD(1,2)`.

## Project file properties worth knowing

Excel-DNA SDK-style build integration recognizes properties such as:

- `RunExcelDnaBuild`
- `RunExcelDnaPack`
- `RunExcelDnaClean`
- `RunExcelDnaSetDebuggerOptions`
- `ExcelDnaPackExePath`
- `ExcelDnaToolsPath`
- `ExcelAddInExplicitExports`

Use these for build customization instead of ad-hoc post-build copy scripts where possible.

## `.dna` file role

Older projects often have explicit `.dna` files. Modern SDK-style projects can generate outputs without a hand-authored `.dna` file for simple cases. Keep or introduce a `.dna` file when the add-in needs explicit external libraries, custom packing options, or legacy behavior.

## Naming strategy

Use a consistent function prefix or namespace. Good UDF libraries avoid collisions with native functions and other add-ins.

Examples:

- `MYCO.PRICE`
- `MYCO.CURVE.BUILD`
- `MYCO.JSON.GET`

Excel's custom function namespace conventions differ between technologies. For Excel-DNA, the function name is the registered Excel name; make it deliberate.

## Architecture checklist

Before writing much code, answer:

- Which Excel versions and bitness are supported?
- Which runtime family is used: .NET Framework, modern .NET, or NativeAOT preview?
- Will the add-in be distributed as a packed `.xll`, installer, or both?
- Is the function library pure UDFs, or does it include UI/commands?
- Are functions thread-safe? Should any be `IsThreadSafe=true`?
- Are async/streaming functions needed?
- Are object handles needed for expensive, stateful, or non-scalar objects?
- How will the add-in be tested inside Excel?
- Will functions be explicitly exported?

## Use the scaffold script

The package includes:

```bash
python scripts/scaffold_excel_dna_addin.py --name MyAddIn --target net48 --output ./work
python scripts/scaffold_excel_dna_addin.py --name MyAddIn --target net8 --features async,ribbon,testing --output ./work
python scripts/scaffold_excel_dna_addin.py --name MyNativeAddIn --target nativeaot --output ./work
```

The script creates a starter project, not a substitute for understanding deployment and runtime choice.
