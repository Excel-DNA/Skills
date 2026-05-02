# NativeAOT Preview for Excel-DNA

## Contents

- [Status and activation](#status-and-activation)
- [What NativeAOT means in .NET](#what-nativeaot-means-in-net)
- [Why NativeAOT matters for Excel-DNA](#why-nativeaot-matters-for-excel-dna)
- [When to recommend NativeAOT](#when-to-recommend-nativeaot)
- [When not to recommend NativeAOT](#when-not-to-recommend-nativeaot)
- [Runtime choice matrix](#runtime-choice-matrix)
- [Minimal NativeAOT preview project](#minimal-nativeaot-preview-project)
- [Build and publish workflow](#build-and-publish-workflow)
- [Package layout and version rules](#package-layout-and-version-rules)
- [Supported Excel-DNA features in the preview](#supported-excel-dna-features-in-the-preview)
- [Features that need caution or are not yet supported](#features-that-need-caution-or-are-not-yet-supported)
- [AOT compatibility rules for add-in code](#aot-compatibility-rules-for-add-in-code)
- [Function design guidance under NativeAOT](#function-design-guidance-under-nativeaot)
- [Async, streaming, and long-running work](#async-streaming-and-long-running-work)
- [Ribbon and Excel object model access](#ribbon-and-excel-object-model-access)
- [Object handles and extended registration](#object-handles-and-extended-registration)
- [Native dependencies and interop](#native-dependencies-and-interop)
- [Testing and release gates](#testing-and-release-gates)
- [Troubleshooting NativeAOT preview projects](#troubleshooting-nativeaot-preview-projects)
- [Docs and product improvement opportunities](#docs-and-product-improvement-opportunities)
- [Agent response patterns](#agent-response-patterns)

## Status and activation

NativeAOT support in Excel-DNA is currently a **preview/specialized path**, not the ordinary default for new add-ins. Use this reference when a user asks about any of these:

- `ExcelDna.AddIn.NativeAOT`
- NativeAOT, AOT, runtime-free XLLs, self-contained XLLs, native compiled add-ins
- avoiding installation of the .NET Desktop Runtime
- 64-bit-only native Excel-DNA distribution
- migrating an existing managed Excel-DNA add-in to AOT
- AOT/trimming warnings such as `IL2026`, `IL3050`, `IL3058`, or missing native code/metadata errors

Default stance: **encourage investigation and prototypes, but avoid promising drop-in compatibility**. Ask the agent to identify the required feature set, dependencies, and deployment constraints before recommending NativeAOT.

## What NativeAOT means in .NET

NativeAOT compiles managed IL into native code during publish. The resulting app is self-contained, does not use a JIT compiler at runtime, targets a specific runtime identifier such as `win-x64`, and can run without the .NET runtime installed.

The trade-off is that the publish process must know what code and metadata will be needed. Code relying on dynamic loading, runtime code generation, unbounded reflection, some COM paths, expression-tree compilation, dynamic serializers, or libraries that are not trimming/AOT-friendly can fail at publish time or runtime.

Important .NET rules to reflect in advice:

- Set `PublishAot` in the project file so build-time analyzers run early, not only in a one-off publish command.
- NativeAOT output is RID-specific; an Excel-DNA NativeAOT preview add-in currently targets 64-bit Excel with `win-x64`.
- Treat AOT and trimming warnings as release blockers unless the project has a documented, tested suppression.
- Libraries intended for AOT should target `net8.0` or later for analyzer coverage. For current preview Excel-DNA NativeAOT packages, expect `net10.0-windows` examples and package metadata.
- NativeAOT can improve startup and remove runtime installation, but it is not automatically faster for every CPU-bound algorithm. Benchmark in Excel.

## Why NativeAOT matters for Excel-DNA

Excel-DNA ordinarily hosts managed code inside Excel through the Excel-DNA loader and the .NET runtime. For modern .NET add-ins this means the .NET Desktop Runtime must be installed and only one modern .NET runtime can be loaded into the Excel process.

NativeAOT changes that runtime story: the add-in's managed code is compiled into native code and packaged as a native 64-bit XLL. This can avoid the modern .NET Desktop Runtime prerequisite and avoid the ordinary modern-.NET in-process runtime collision problem.

That makes NativeAOT strategically important for Excel-DNA because it can combine:

- the native `.xll` deployment shape Excel already understands
- C#/.NET source code and Excel-DNA function registration
- no end-user .NET Desktop Runtime installation
- no JIT startup cost
- a single 64-bit XLL distribution artifact for compatible add-ins

It does **not** make the add-in cross-platform. XLLs remain Excel-for-Windows add-ins.

## When to recommend NativeAOT

Recommend a NativeAOT investigation when most of these are true:

- The add-in targets **64-bit Excel for Windows only**.
- Runtime-free distribution is a high-value requirement.
- The feature set is mostly UDFs, commands, simple ribbon, object handles, and Excel-DNA-supported extended registration features that have been tested in the preview.
- Dependencies are simple, statically referenced, trimming/AOT-compatible, and do not depend on dynamic plugin loading.
- The team can run `dotnet publish` and Excel smoke tests often.
- The team accepts preview risk and can pin package versions.
- The release process can treat AOT warnings as defects.

Good early pilots:

- computational UDF libraries with stable dependencies
- internal 64-bit-only add-ins where preview risk is acceptable
- distribution-sensitive add-ins where installing the Desktop Runtime is the main blocker
- small add-ins used to benchmark NativeAOT startup/load behavior and formula performance

## When not to recommend NativeAOT

Avoid NativeAOT, or keep it as a parallel experiment, when any of these are true:

- The add-in must support 32-bit Excel.
- The add-in relies on `ExcelDna.IntelliSense` overlay UI, custom task panes, rich WinForms/WPF UI, or other UI surfaces not yet proven under the NativeAOT preview.
- The add-in loads plugins or assemblies dynamically.
- The add-in depends heavily on reflection, dynamic proxies, expression compilation, runtime code generation, serializers without source generation, or C++/CLI.
- The team needs stable, production-ready package/tooling behavior immediately.
- The add-in already works well on `.NET Framework 4.8` and the main distribution pain is Trust Center/signing rather than runtime installation.
- The add-in targets cross-platform Excel; Office.js is the cross-platform family, not XLL/NativeAOT.

## Runtime choice matrix

| Runtime path | Best fit | Main risk |
|---|---|---|
| `.NET Framework 4.8` | broad external Windows distribution; stable production default | older .NET and C# ecosystem |
| Modern .NET framework-dependent | controlled desktops needing newer .NET/C# | Desktop Runtime prerequisite and one modern runtime per Excel process |
| NativeAOT preview | runtime-free 64-bit native XLL distribution for compatible add-ins | preview package/tooling, AOT/trimming restrictions, feature gaps |
| Native C/C++ XLL | maximum low-level control and native codebase | C/C++ complexity and C API burden |

## Minimal NativeAOT preview project

Use a separate project, or carefully conditional project file sections, rather than unconditionally mixing ordinary managed and NativeAOT package references.

As of the first NativeAOT preview train, examples use `net10.0-windows`, `win-x64`, `PublishAot`, and `ExcelDna.AddIn.NativeAOT`. Use the current matching preview package version from the feed you are targeting.

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

Minimal UDF:

```csharp
using ExcelDna.Integration;

namespace MyNativeAddIn;

public static class Functions
{
    [ExcelFunction(
        Name = "NATIVE.HELLO",
        Description = "Returns a greeting from a NativeAOT Excel-DNA add-in.",
        Category = "NativeAOT Preview")]
    public static string Hello(
        [ExcelArgument(Name = "name", Description = "Name to greet")] string name)
    {
        return $"Hello from NativeAOT, {name}!";
    }

    [ExcelFunction(
        Name = "NATIVE.ADD",
        Description = "Adds two numbers.",
        Category = "NativeAOT Preview",
        IsThreadSafe = true)]
    public static double Add(double x, double y) => x + y;
}
```

Use `using ExcelDna.Integration;` even when the only explicit package reference is `ExcelDna.AddIn.NativeAOT`; the preview package should bring the integration types for the simple project layout.

## Build and publish workflow

NativeAOT builds need the .NET SDK and native toolchain. On Windows, install a Visual Studio version supported by the .NET SDK with the Desktop development with C++ workload.

Recommended commands:

```bash
dotnet restore -r win-x64
dotnet publish -c Release -r win-x64
```

Expected output shape is RID-specific. The published XLL is typically under a path similar to:

```text
bin/Release/net10.0-windows/win-x64/publish/win-x64/MyNativeAddIn-AddIn64.xll
```

The `*-AddIn64.xll` is the add-in artifact to load in 64-bit Excel. Build systems should copy/sign/package that file, not the ordinary managed build output.

## Package layout and version rules

For current preview packages:

- Prefer `ExcelDna.AddIn.NativeAOT` alone in a NativeAOT project.
- Do not unconditionally reference both `ExcelDna.AddIn` and `ExcelDna.AddIn.NativeAOT` in the same project.
- If a multi-target or conditional setup needs multiple Excel-DNA packages, keep all `ExcelDna.*` versions aligned on the same preview train.
- If a preview/custom setup fails because `ExcelDnaToolsPath` points at the NativeAOT tools folder but base loader assets are expected elsewhere, use the documented workaround only for that preview layout:

```xml
<PropertyGroup>
  <ExcelDnaToolsPath>$(PkgExcelDna_AddIn)/tools/</ExcelDnaToolsPath>
</PropertyGroup>
```

Prefer avoiding the workaround by using a NativeAOT-only project until the preview package layout stabilizes.

## Supported Excel-DNA features in the preview

Based on the current Excel-DNA NativeAOT page and public preview discussion, the preview has examples for:

- scalar UDFs with `[ExcelFunction]`
- async functions via `[ExcelAsyncFunction]` and task-returning functions
- `IExcelAddIn` lifecycle hooks such as `AutoOpen` and `AutoClose`
- `[ExcelCommand]` commands
- simple ribbon via the NativeAOT `IExcelRibbon` pattern
- Excel object model access through `ExcelDnaUtil.DynamicApplication` and `IDynamic`
- Function Wizard metadata through `[ExcelFunction]` and `[ExcelArgument]`
- range/array parameters such as `object[,]`
- nullable and optional parameters
- range parameters through `IRange`
- enum parameters and return values
- string arrays and 2D string arrays
- limited `params` arguments
- object handles via `[ExcelHandle]` and external handle registration
- user-defined parameter and return conversions
- function execution handlers

Treat this list as **feature candidates that still need project-specific tests**, not a blanket production certification.

## Features that need caution or are not yet supported

Do not promise these in NativeAOT preview unless the current source/tests prove them:

- `ExcelDna.IntelliSense` overlay UI. Public discussion indicates it does not work yet under NativeAOT preview and depends on further WinForms/trim compatibility work.
- Custom task panes and rich WinForms/WPF UI. Treat as unproven unless the specific UI stack is tested.
- Ribbon images. The current NativeAOT support page lists loading images for ribbon controls as unsupported.
- Diagnostic Display clipboard copy. The current support page lists this as unsupported.
- Dynamic assembly loading, runtime-generated wrappers, and plugin systems.
- C++/CLI dependencies.
- Traditional COM interop assumptions. NativeAOT lacks built-in Windows COM support in the same way ordinary .NET desktop code uses it; Excel-DNA's `DynamicApplication`/`IDynamic` path is the preview-safe direction.
- Arbitrary expression-tree compilation. Expression trees use interpreted form under NativeAOT and can be slower or require metadata/code that is not generated.

## AOT compatibility rules for add-in code

Review the project as a closed-world native build:

- All code needed at runtime must be statically discoverable or explicitly preserved.
- Avoid `Assembly.Load`, `Assembly.LoadFile`, plugin discovery by reflection, dynamic proxy generation, `Reflection.Emit`, and broad `Type.GetType`/member reflection patterns.
- Prefer source-generated serializers over reflection-heavy serializers.
- Prefer typed wrappers and explicit conversion functions over runtime-built expression wrappers.
- Avoid dependencies that emit `IL2026`, `IL3050`, or related warnings unless the warning is understood, documented, and covered by Excel smoke tests.
- Enable AOT analysis early. Consider library settings such as `IsAotCompatible` for shared libraries that must support AOT consumers.
- Publish frequently during development. A project can compile normally but fail or warn only during NativeAOT publish.

A NativeAOT project is only healthy when the publish is warning-clean or every warning has a tracked justification and runtime test.

## Function design guidance under NativeAOT

Prefer functions that are:

- static, explicitly attributed, and simple to discover
- strongly typed with Excel-friendly parameter and return types
- free of reflection-heavy conversion logic
- deterministic and thread-safe where marked
- explicit about array/range shapes
- explicit about errors, missing values, and optional arguments

For high-performance compute:

- Benchmark managed modern .NET, `.NET Framework`, and NativeAOT inside Excel, not just in a console app.
- Separate Excel marshaling overhead from algorithm time.
- Compare cold load time, first calculation, repeated calculation, memory footprint, and binary size.
- Remember that JIT/tiered/PGO optimizations can make framework-dependent modern .NET faster for some long-running workloads.

## Async, streaming, and long-running work

NativeAOT preview examples include async functions and task-returning functions. Still apply the general Excel-DNA async rules:

- Use async for I/O-bound work or work that should not block Excel.
- Do not mutate the workbook from background work.
- Include all function inputs in the function identity/signature.
- Add cancellation and timeouts when supported.
- Treat `System.Linq.Expressions` or generated delegate wrappers with caution because NativeAOT cannot generate new native code at runtime.

If an async NativeAOT add-in fails with missing native code or metadata for a large generic `Expression<Func<...>>`, reduce the example to a minimal function and check whether the failure comes from a registration wrapper, a custom conversion path, or a dependency that is not AOT-compatible.

## Ribbon and Excel object model access

For ordinary managed Excel-DNA, ribbon classes usually derive from `ExcelRibbon`. For NativeAOT preview, use the NativeAOT docs/sample pattern first:

```csharp
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;

namespace MyNativeAddIn;

public class RibbonController : IExcelRibbon
{
    public string GetCustomUI(string ribbonId)
    {
        return @"
<customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui'>
  <ribbon>
    <tabs>
      <tab id='NativeTab' label='NativeAOT'>
        <group id='NativeGroup' label='Tools'>
          <button id='WriteFormulaButton' label='Write Formula' onAction='OnWriteFormula'/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
    }

    public void OnWriteFormula(RibbonControl control)
    {
        IDynamic app = ExcelDnaUtil.DynamicApplication;
        app.Get<IDynamic>("ActiveCell").Set("Formula", "=NATIVE.ADD(1,2)");
    }
}
```

Prefer `ExcelDnaUtil.DynamicApplication` and `IDynamic` in NativeAOT preview code. Avoid assuming that C# `dynamic`, VSTO PIAs, or ordinary COM interop patterns behave the same as in framework-dependent desktop .NET.

## Object handles and extended registration

Object handles are a promising NativeAOT use case because they let Excel formulas work with complex .NET objects while keeping the worksheet surface small.

NativeAOT preview examples include:

- `[return: ExcelHandle]` on factory functions
- `[ExcelHandle]` on classes
- `[assembly: ExcelHandleExternal(typeof(SomeType))]` for external handle types
- user-defined parameter and return conversions
- function execution handlers

Guidance:

- Keep handle types simple and statically known.
- Avoid dynamic proxy or reflection-heavy handle factories.
- Provide inspection functions so users can see handle type/status.
- Provide lifecycle and cleanup policies.
- Test recalculation, workbook close, add-in unload, and repeated create/destroy cycles.

## Native dependencies and interop

NativeAOT can improve interop performance and supports P/Invoke, but native dependency handling becomes a native-linking/deployment topic:

- P/Invokes are normally lazily bound; direct P/Invoke can be configured for selected entry points when startup binding and performance matter.
- Direct P/Invoke changes dependency search behavior and requires the target library/entry point at load time.
- Static linking can be configured with native libraries where appropriate.
- The final XLL must still match Excel bitness and Windows architecture.
- Test native dependencies on clean machines. A developer machine often hides missing DLLs or PATH assumptions.

## Testing and release gates

A NativeAOT preview add-in should not pass release review until these checks succeed:

1. `dotnet publish -c Release -r win-x64` completes.
2. NativeAOT/trimming warnings are zero, or every warning has a linked justification and targeted test.
3. The published `*-AddIn64.xll` loads in 64-bit Excel on a clean machine without the .NET Desktop Runtime installed.
4. Core UDFs calculate correctly.
5. Async/streaming/object-handle/ribbon features are tested if present.
6. `ExcelDna.IntelliSense` is not included unless current preview support has been proven.
7. The add-in is signed and Trust Center/MOTW behavior has been tested.
8. The package version and preview status are visible in release notes.
9. Benchmark claims include cold load, first calculation, repeated calculation, and memory footprint.
10. A fallback managed build exists if NativeAOT is experimental for a production customer.

## Troubleshooting NativeAOT preview projects

| Symptom | Likely cause | Response |
|---|---|---|
| `ExcelFunction` attribute not found | missing `using ExcelDna.Integration` or wrong package layout | start from the minimal NativeAOT-only project |
| `Loading ExcelDna.ManagedHost failed` | mixed package references, wrong output, preview packaging issue, or trying to load non-published artifact | use NativeAOT package only, publish, load `*-AddIn64.xll` from publish output |
| `File does not exist (Xll32FilePath)` or missing `ExcelDna.xll` | preview `ExcelDnaToolsPath` conflict when packages are mixed | avoid mixing packages or apply the temporary tools-path workaround |
| AOT warning `IL3050` | runtime code generation or `RequiresDynamicCode` API | remove/replace dynamic pattern, use source generation/static code, retest |
| Trim warning `IL2026` | code may use members removed by trimming | replace API, annotate/preserve deliberately, or avoid NativeAOT |
| Missing native code/metadata at runtime | publish warnings ignored, reflection/dynamic generic path not preserved | reduce to minimal repro, fix warnings, add targeted test |
| IntelliSense overlay not working | not supported in preview | use metadata/Function Wizard; avoid `ExcelDna.IntelliSense` package |
| Ribbon image missing | unsupported preview feature | use text-only controls or defer image support |
| Works on developer machine only | wrong artifact, unsigned/blocked XLL, missing native dependency, policy issue | test signed publish output on clean 64-bit Excel machine |

## Docs and product improvement opportunities

NativeAOT is important enough to deserve first-class docs and tooling even while preview:

- Add a top-level runtime decision table with NativeAOT as a separate path.
- Add a current NativeAOT quickstart that pins the preview package version and says exactly which file to load.
- Document the ordinary managed add-in path versus NativeAOT path side by side.
- Clearly state that `ExcelDna.IntelliSense` overlay and CTPs are not yet normal NativeAOT features.
- Add a feature support matrix with `supported`, `preview-supported`, `unsupported`, and `unknown/needs test` states.
- Add a NativeAOT migration checklist for existing 1.9 add-ins.
- Add a diagnostic page for AOT warnings and common runtime errors.
- Add `dotnet new` NativeAOT preview template.
- Add project analyzer warnings for mixed `ExcelDna.AddIn`/`ExcelDna.AddIn.NativeAOT`, missing `PublishAot`, missing `win-x64`, and accidental `ExcelDna.IntelliSense` references.
- Add CI recipes for publish and clean-machine Excel smoke tests.

## Agent response patterns

When the user asks, "Should I use NativeAOT?", answer with:

1. Preview status.
2. Runtime-free 64-bit benefit.
3. Feature/dependency compatibility check.
4. Minimal pilot project.
5. Publish and clean-machine test plan.
6. Fallback recommendation if they need stable production support now.

When the user asks for a new NativeAOT add-in, generate:

- NativeAOT-only project file.
- `net10.0-windows`, `win-x64`, `PublishAot`.
- `ExcelDna.AddIn.NativeAOT` preview package.
- Explicitly attributed UDF.
- Publish command and exact output artifact guidance.
- Warning that `ExcelDna.IntelliSense` and CTPs are not part of the safe preview baseline.

When the user asks to migrate an existing add-in, do not just change the package. First inventory:

- UI surfaces.
- IntelliSense/CTP usage.
- async/streaming patterns.
- registration transforms and execution handlers.
- object handles.
- dependencies and native dependencies.
- reflection/dynamic/runtime-codegen usage.
- 32-bit support requirements.
