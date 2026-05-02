#!/usr/bin/env python3
"""Scaffold a small Excel-DNA add-in project.

This creates starter files only; it does not install packages, run dotnet, publish NativeAOT, or modify Excel.
"""
from __future__ import annotations

import argparse
import re
from pathlib import Path

PACKAGE_VERSION = "1.9.0"
NATIVE_AOT_PACKAGE_VERSION = "1.10.0-preview4"


def safe_identifier(name: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9_]", "", name)
    if not cleaned or cleaned[0].isdigit():
        cleaned = "MyAddIn" + cleaned
    return cleaned


def target_framework(target: str) -> tuple[str, str]:
    if target == "net48":
        return "net48", ""
    if target == "net10":
        return "net10.0-windows", "<UseWindowsForms>true</UseWindowsForms>\n    <RollForward>LatestMajor</RollForward>"
    if target == "net8":
        return "net8.0-windows", "<UseWindowsForms>true</UseWindowsForms>\n    <RollForward>LatestMajor</RollForward>"
    if target == "multi":
        return "net48;net10.0-windows", "<UseWindowsForms>true</UseWindowsForms>"
    if target == "nativeaot":
        return "net10.0-windows", "<RuntimeIdentifier>win-x64</RuntimeIdentifier>\n    <PublishAot>true</PublishAot>"
    raise ValueError(f"Unsupported target: {target}")


def write_project(out: Path, name: str, target: str, features: set[str]) -> None:
    ident = safe_identifier(name)
    native_aot = target == "nativeaot"
    tfm, extra = target_framework(target)
    tfm_tag = "TargetFrameworks" if target == "multi" else "TargetFramework"
    extra_block = f"\n    {extra}" if extra else ""
    if native_aot:
        pack_props = ""
        package_line = f'<PackageReference Include="ExcelDna.AddIn.NativeAOT" Version="{NATIVE_AOT_PACKAGE_VERSION}" />'
    else:
        pack_props = "\n    <RunExcelDnaPack>true</RunExcelDnaPack>"
        package_line = f'<PackageReference Include="ExcelDna.AddIn" Version="{PACKAGE_VERSION}" />'

    csproj = f"""
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <{tfm_tag}>{tfm}</{tfm_tag}>{extra_block}
    <Nullable>enable</Nullable>
    <ImplicitUsings>enable</ImplicitUsings>
    <LangVersion>latest</LangVersion>
    <ExcelAddInExplicitExports>true</ExcelAddInExplicitExports>{pack_props}
  </PropertyGroup>

  <ItemGroup>
    {package_line}
  </ItemGroup>
</Project>
""".lstrip()
    (out / f"{ident}.csproj").write_text(csproj, encoding="utf-8")

    category = "NativeAOT Preview" if native_aot else ident
    functions = f"""
using ExcelDna.Integration;

namespace {ident};

public static class Functions
{{
    [ExcelFunction(Name = "{ident.upper()}.ADD", Description = "Adds two numbers.", Category = "{category}", IsThreadSafe = true)]
    public static double Add(
        [ExcelArgument(Name = "x", Description = "First number")] double x,
        [ExcelArgument(Name = "y", Description = "Second number")] double y)
        => x + y;

    [ExcelFunction(Name = "{ident.upper()}.HELLO", Description = "Returns a greeting.", Category = "{category}")]
    public static string Hello(
        [ExcelArgument(Name = "name", Description = "Name to greet")] string name)
        => $"Hello, {{name}}";
}}
""".lstrip()
    (out / "Functions.cs").write_text(functions, encoding="utf-8")

    if "async" in features:
        if native_aot:
            async_code = f"""
using ExcelDna.Integration;
using System.Threading;
using System.Threading.Tasks;

namespace {ident};

public static class AsyncFunctions
{{
    [ExcelFunction(Name = "{ident.upper()}.DELAYED.ADD", Description = "Adds two numbers after an asynchronous delay.", Category = "NativeAOT Preview")]
    public static async Task<object> DelayedAdd(double x, double y, CancellationToken cancellationToken)
    {{
        try
        {{
            await Task.Delay(250, cancellationToken).ConfigureAwait(false);
            return x + y;
        }}
        catch (OperationCanceledException)
        {{
            return ExcelError.ExcelErrorNA;
        }}
    }}
}}
""".lstrip()
        else:
            async_code = f"""
using ExcelDna.Integration;
using System;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;

namespace {ident};

public static class AsyncFunctions
{{
    private static readonly HttpClient Http = new() {{ Timeout = TimeSpan.FromSeconds(15) }};

    [ExcelFunction(Name = "{ident.upper()}.HTTP.GET", Description = "Gets text from a URL asynchronously.", Category = "{ident}")]
    public static async Task<object> HttpGet(string url, CancellationToken cancellationToken)
    {{
        try
        {{
            using var response = await Http.GetAsync(url, cancellationToken).ConfigureAwait(false);
            response.EnsureSuccessStatusCode();
            return await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
        }}
        catch (OperationCanceledException)
        {{
            return ExcelError.ExcelErrorNA;
        }}
        catch
        {{
            return ExcelError.ExcelErrorValue;
        }}
    }}
}}
""".lstrip()
        (out / "AsyncFunctions.cs").write_text(async_code, encoding="utf-8")

    if "ribbon" in features:
        if native_aot:
            ribbon_code = f"""
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;

namespace {ident};

public class RibbonController : IExcelRibbon
{{
    public string GetCustomUI(string ribbonId)
    {{
        return @"
<customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui'>
  <ribbon>
    <tabs>
      <tab id='{ident}Tab' label='{ident}'>
        <group id='ToolsGroup' label='Tools'>
          <button id='InsertFormulaButton' label='Insert Formula' onAction='OnInsertFormula' />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
    }}

    public void OnInsertFormula(RibbonControl control)
    {{
        IDynamic app = ExcelDnaUtil.DynamicApplication;
        app.Get<IDynamic>("ActiveCell").Set("Formula", "={ident.upper()}.ADD(1,2)");
    }}
}}
""".lstrip()
        else:
            ribbon_code = f"""
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using System.Runtime.InteropServices;

namespace {ident};

[ComVisible(true)]
public class RibbonController : ExcelRibbon
{{
    public override string GetCustomUI(string ribbonId)
    {{
        return @"
<customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui'>
  <ribbon>
    <tabs>
      <tab id='{ident}Tab' label='{ident}'>
        <group id='ToolsGroup' label='Tools'>
          <button id='InsertFormulaButton' label='Insert Formula' onAction='OnInsertFormula' />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
    }}

    public void OnInsertFormula(IRibbonControl control)
    {{
        dynamic excel = ExcelDnaUtil.Application;
        excel.ActiveCell.Formula = "={ident.upper()}.ADD(1,2)";
    }}
}}
""".lstrip()
        (out / "RibbonController.cs").write_text(ribbon_code, encoding="utf-8")

    if native_aot:
        readme_notes = f"""
- Target selection: `nativeaot` preview.
- The project references `ExcelDna.AddIn.NativeAOT`, not ordinary `ExcelDna.AddIn`.
- It targets 64-bit Excel with `net10.0-windows`, `win-x64`, and `PublishAot`.
- Publish with `dotnet publish -c Release -r win-x64` and load the published `*-AddIn64.xll` from the publish folder.
- Treat NativeAOT/trimming warnings as defects until investigated.
- `ExcelDna.IntelliSense`, custom task panes, ribbon images, and rich COM assumptions are not part of the safe preview baseline.
""".strip()
    else:
        readme_notes = f"""
- Target selection: `{target}`.
- Explicit exports are enabled.
- Build the project and load the generated `.xll` in Excel for Windows.
""".strip()
        if "testing" in features:
            readme_notes += "\n- For Excel-hosted tests, create a separate test project or harness that references `ExcelDna.Testing`; do not add the test harness package to this add-in project by default."

    readme = f"""
# {ident}

Excel-DNA starter add-in.

## Test formulas

```excel
={ident.upper()}.ADD(1,2)
={ident.upper()}.HELLO("Excel")
```

## Notes

{readme_notes}
""".lstrip()
    (out / "README.md").write_text(readme, encoding="utf-8")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--name", required=True, help="Project/add-in name")
    parser.add_argument("--target", choices=["net48", "net10", "net8", "multi", "nativeaot"], default="net48")
    parser.add_argument("--features", default="", help="Comma-separated: async,ribbon,testing")
    parser.add_argument("--output", default=".", help="Output directory")
    args = parser.parse_args()

    features = {f.strip().lower() for f in args.features.split(",") if f.strip()}
    allowed = {"async", "ribbon", "testing"}
    unknown = features - allowed
    if unknown:
        raise SystemExit(f"Unknown feature(s): {', '.join(sorted(unknown))}")

    project_dir = Path(args.output).resolve() / safe_identifier(args.name)
    project_dir.mkdir(parents=True, exist_ok=True)
    write_project(project_dir, args.name, args.target, features)
    print(project_dir)


if __name__ == "__main__":
    main()
