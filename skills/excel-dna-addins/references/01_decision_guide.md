# Decision Guide: Excel-DNA, XLLs, Office.js, VBA, VSTO, C/C++, and Runtime Choice

## Contents

- [When Excel-DNA is the strong choice](#when-excel-dna-is-the-strong-choice)
- [When Excel-DNA is not the best default](#when-excel-dna-is-not-the-best-default)
- [Excel-DNA vs Office.js custom functions](#excel-dna-vs-officejs-custom-functions)
- [Excel-DNA vs native C/C++ XLL](#excel-dna-vs-native-cc-xll)
- [Excel-DNA vs VBA](#excel-dna-vs-vba)
- [Excel-DNA vs VSTO / COM add-ins](#excel-dna-vs-vsto-com-add-ins)
- [Runtime choice: .NET Framework vs modern .NET](#runtime-choice-net-framework-vs-modern-net)
  - [.NET Framework 4.8 / 4.8.1](#net-framework-48-481)
  - [Modern .NET (`net8.0-windows` and later where supported)](#modern-net-net80-windows-and-later-where-supported)
  - [NativeAOT preview](#nativeaot-preview)
- [Recommended default choices](#recommended-default-choices)
- [Product positioning summary](#product-positioning-summary)

## When Excel-DNA is the strong choice

Excel-DNA is strongest when the product is an Excel-on-Windows add-in whose main value is exposing fast, maintainable .NET code as worksheet functions. It is also strong when those functions must coexist with Excel-native calculation, arrays, recalculation, RTD/streaming, ribbons, task panes, COM access, and ordinary Windows desktop deployment.

Use Excel-DNA when:

- The primary audience uses Excel for Windows.
- The add-in must expose user-defined functions with natural worksheet formula behavior.
- You want C#, F#, or VB.NET rather than C/C++.
- You need direct or indirect access to Excel's C API and XLL registration model.
- You need a single `.xll` distribution artifact or a Windows installer.
- You want high-performance local computation without web-service latency.
- You want to leverage existing .NET libraries, corporate code, or quantitative/engineering libraries.

## When Excel-DNA is not the best default

Choose another route when:

- Cross-platform Excel support is the main requirement: web, Mac, iPad, and Windows. Office.js is usually the right add-in family.
- The add-in is mostly a workbook automation macro for a small internal audience. VBA may be faster to author and easier for power users to inspect.
- The add-in must integrate deeply with Office extensibility beyond Excel or needs classic COM add-in surfaces across Office applications. VSTO or COM may be relevant, though VSTO has distribution and platform constraints.
- The add-in is a native C/C++ numerical XLL and every microsecond matters, or the team already owns a native XLL pipeline. Excel-DNA can still be a wrapper or migration path.

## Excel-DNA vs Office.js custom functions

Office.js add-ins use web technologies and are designed for cross-platform Office. Their custom functions can appear in worksheets and can stream web-originated data. The trade-off is that the runtime, APIs, and deployment model are web add-in centric. Office.js does not expose all XLL capabilities, and Microsoft itself notes that an XLL may be the better choice for some Windows-specific functionality.

Excel-DNA add-ins are `.xll` add-ins, so they are Windows-only. In return, they fit Excel's native add-in and calculation ecosystem closely, can use .NET libraries locally, can integrate with Excel's C API, and can support high-performance local functions without requiring a web runtime.

Decision heuristic:

| Requirement | Prefer |
|---|---|
| Windows-only power users, finance/quant/engineering models, local .NET computation | Excel-DNA |
| Cross-platform including web/Mac/iPad | Office.js |
| Existing XLL to modernize while keeping Windows semantics | Excel-DNA or C/C++ XLL |
| Internal macro automation only | VBA or Excel-DNA command/ribbon depending maintainability |
| Need .NET libraries and worksheet UDFs | Excel-DNA |
| Need web-service orchestration and centralized deployment across platforms | Office.js |

## Excel-DNA vs native C/C++ XLL

Native XLLs are the closest possible path to the Excel C API. They can be extremely fast and precise but require C/C++ and manual care around memory, registration, threading, and API constraints.

Excel-DNA makes the XLL model practical from .NET by hosting managed code, handling marshaling and registration, and exposing `XlCall`/`ExcelReference` when needed. For most teams, Excel-DNA gives a better productivity/performance trade-off than writing a complete C/C++ XLL.

Use native C/C++ when:

- The algorithm already exists in native code and wrapping is simpler than porting.
- Latency is dominated by language/runtime overhead and .NET is unacceptable.
- You need exact control over C API memory and registration.

Use Excel-DNA when:

- .NET maintainability and ecosystem matter.
- UDF metadata and registration should be generated from attributes.
- The add-in includes async, streaming, ribbons, CTPs, IntelliSense, or installer packaging.

## Excel-DNA vs VBA

VBA remains excellent for workbook-local automation and user-maintained macros. Excel-DNA is stronger for distributable add-ins, typed code, source control, testing, NuGet dependencies, and high-quality UDF libraries.

A common migration path is:

1. Keep the workbook model and existing formulas.
2. Replace performance-critical or reusable VBA UDFs with Excel-DNA functions.
3. Move UI actions to a ribbon or commands.
4. Use tests around both Excel models and .NET code.

## Excel-DNA vs VSTO / COM add-ins

VSTO and COM add-ins are centered on Office automation, UI, and application lifecycle. They are not the natural choice for high-volume worksheet UDF libraries. Excel-DNA can expose COM-visible classes, ribbon UI, and task panes while keeping UDFs first-class.

## Runtime choice: .NET Framework vs modern .NET

### .NET Framework 4.8 / 4.8.1

Advantages:

- Installed as part of supported Windows environments or easy to rely on in enterprises.
- Stable, conservative, and unlikely to force end-user runtime decisions.
- Separate AppDomain isolation between add-ins is stronger than modern .NET in-process isolation.
- Best default for broad external distribution where target machines are uncontrolled.

Trade-offs:

- No new C# runtime features.
- Older BCL and NuGet compatibility constraints.
- Less attractive for new .NET ecosystem libraries.

### Modern .NET (`net8.0-windows` and later where supported)

Advantages:

- Newer C# language features and .NET libraries.
- Better alignment with current .NET development.
- Potential NativeAOT route for specialized scenarios.

Trade-offs:

- Requires compatible .NET Desktop Runtime on the user's machine unless NativeAOT is used.
- Only one modern .NET runtime can be loaded into the Excel process; add-ins targeting incompatible runtime versions can collide.
- Runtime support cadence matters; avoid out-of-support targets.
- Weak add-in isolation compared with .NET Framework AppDomains.

### NativeAOT preview

NativeAOT is a preview/specialized route for producing native 64-bit Excel-DNA XLLs that do not require the .NET Desktop Runtime to be installed on user machines. It compiles managed IL to native code at publish time and targets a specific runtime identifier such as `win-x64`.

Advantages:

- Runtime-free distribution for compatible 64-bit Excel add-ins.
- Potentially faster cold startup and lower memory footprint.
- Avoids the ordinary modern-.NET Desktop Runtime prerequisite.
- Avoids the ordinary one-modern-.NET-runtime-per-Excel-process collision story because the add-in is native compiled.

Trade-offs:

- Preview package/tooling behavior; do not treat as stable production default yet.
- Current Excel-DNA preview path is 64-bit-focused.
- NativeAOT requires trimming/AOT compatibility; dynamic loading, runtime code generation, some reflection, C++/CLI, and ordinary COM assumptions are restricted.
- Some Excel-DNA surfaces are unsupported or unproven in the preview, especially `ExcelDna.IntelliSense` overlay UI, CTPs, ribbon images, and rich WinForms/WPF UI.
- Current examples use `ExcelDna.AddIn.NativeAOT`, `net10.0-windows`, `RuntimeIdentifier=win-x64`, and `PublishAot=true`; package versions must be kept aligned on the preview train.

Use NativeAOT as a deliberate compatibility project: create a small pilot, publish often, treat AOT warnings as defects, test in 64-bit Excel on a clean machine, and keep a managed fallback when production stability matters. Open `references/11_native_aot_preview.md` before giving detailed setup or migration advice.

## Recommended default choices

| Scenario | Default recommendation |
|---|---|
| Public/external add-in with unknown target machines | `net48`, packed `.xll`, code-signing, installer optional |
| Corporate add-in with managed desktops | `net8.0-windows` or current supported modern .NET, explicit runtime prerequisite |
| Library must support both old and modern deployments | Multi-target `net48;net8.0-windows` |
| Runtime-free 64-bit distribution is strategic and feature set is compatible | Investigate NativeAOT preview as a separate track |
| Experimental sample/tutorial | `net48` for broadest frictionless loading, current supported modern .NET for controlled desktops, or NativeAOT preview only when the exercise is specifically about runtime-free 64-bit XLLs |

## Product positioning summary

Excel-DNA is not merely a way to automate Excel. It is a way to make Excel a typed, testable, distributable front end for .NET computation while preserving the worksheet as the user's modeling language.
