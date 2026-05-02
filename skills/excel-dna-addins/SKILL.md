---
name: excel-dna-addins
description: Use this skill when creating, reviewing, extending, testing, troubleshooting, or distributing Excel-DNA Excel .xll add-ins for Microsoft Excel on Windows using .NET. Covers technology choice, UDF-first design, function registration, async/streaming/RTD, object handles, ribbons, custom task panes, IntelliSense, COM and C API interop, testing, packing, signing, installation, NativeAOT preview builds, WiX installers, and modernization/documentation gaps. Do not use for Office.js-only add-ins except for comparison, migration, or hybrid recommendations.
---

# Excel-DNA Add-ins Skill

This skill is a high-detail working corpus for agents helping users build Excel add-ins with Excel-DNA. It is organized for progressive disclosure: start here, open only the references needed for the user's task, and use the templates/scripts when generating code.

## Activation boundaries

Use this skill for:

- Choosing whether Excel-DNA is the right add-in technology.
- Creating or maintaining `.xll` add-ins for Excel on Windows with C#, F#, or VB.NET.
- Designing Excel user-defined functions (UDFs) as the core add-in surface.
- Implementing advanced functions: thread-safe, macro-sheet, async, streaming, object handles, array/range functions, optional/default parameters, and aspect-like registration transforms.
- Adding Excel extension surfaces supported by Excel-DNA: ribbon UI, COM automation exposure, custom task panes, and function IntelliSense.
- Testing with Excel-hosted helpers and ordinary .NET tests.
- Packaging, signing, installation, runtime selection, NativeAOT preview evaluation, and WiX-based installers.
- Reasoning about Excel's C API, `XlCall`, `ExcelReference`, and add-in registration rules as they affect Excel-DNA add-ins.

Avoid this skill for:

- Office.js/JavaScript-only add-ins with no Excel-DNA comparison or interop question.
- Mac, web, or iPad add-in implementation details where a `.xll` cannot be loaded.
- General .NET application development unrelated to Excel add-ins.

## Quick route map

1. **Technology choice or roadmap**: open `references/01_decision_guide.md`.
2. **New add-in bootstrap**: open `references/02_project_bootstrap.md`; optionally run `scripts/scaffold_excel_dna_addin.py`.
3. **UDF implementation**: open `references/03_udf_core_patterns.md`.
4. **Registration, metadata, conversions, AOP-like wrapping**: open `references/04_function_registration_extended.md`.
5. **Async, streaming, RTD, cancellation, batching**: open `references/05_async_streaming_rtd.md`.
6. **Ribbon, CTP, IntelliSense, COM callbacks**: open `references/06_ribbon_ctp_intellisense.md`.
7. **C API and COM interop**: open `references/07_com_and_c_api.md`.
8. **Testing and distribution**: open `references/08_testing_distribution_installation.md`.
9. **NativeAOT preview builds**: open `references/11_native_aot_preview.md` before giving setup, migration, feature-compatibility, or troubleshooting advice for `ExcelDna.AddIn.NativeAOT`.
10. **Troubleshooting**: open `references/09_troubleshooting_and_known_issues.md`.
11. **Documentation holes and modernization ideas**: open `references/10_modernisation_gaps_improvements.md` and `outputs/modernisation_backlog.md`.

## Golden rules for Excel-DNA work

- Treat the UDF surface as the primary product. Most Excel-DNA value comes from exposing .NET computation as worksheet functions that participate naturally in Excel's recalculation, dependency, auditing, and modeling workflow.
- Prefer simple, deterministic, side-effect-free worksheet functions unless the requirement explicitly needs async, streaming, object handles, macros, or UI interaction.
- Keep UI actions and workbook mutations out of normal worksheet functions. Put state-changing work in ribbon callbacks, commands, macros, or code queued with `ExcelAsyncUtil.QueueAsMacro`.
- Decide runtime and deployment up front. `.NET Framework 4.8` is the conservative distribution default for broad Windows fleets. Modern .NET examples should target `net10.0-windows` as the current LTS baseline unless the user has a specific older supported runtime constraint. Modern .NET gives newer C# and library access but requires installed runtimes and has process-level runtime collision considerations. NativeAOT is a preview/specialized runtime-free 64-bit path; never treat it as a drop-in default without checking feature compatibility and AOT warnings.
- Be explicit about registration. For production libraries, prefer `[ExcelFunction]` and `ExcelAddInExplicitExports=true` unless the add-in is intentionally a quick prototype.
- Respect Excel threading. Never call Excel's COM object model from arbitrary background threads. Use macro/ribbon contexts or queue work back to Excel.
- Use Excel-DNA build tasks and packing instead of hand-copying outputs. Sign packed `.xll` files for distribution.
- For complex behavior, create a small sample workbook and a smoke-test path early.

## Default output style for generated code

When generating an Excel-DNA project:

- Use SDK-style `.csproj`.
- Include `PackageReference Include="ExcelDna.AddIn" Version="1.9.0"` for ordinary managed add-ins unless the user asks to float or pin another version. For NativeAOT preview projects, use `ExcelDna.AddIn.NativeAOT` from the matching preview train instead.
- Prefer `net48` for broad distribution stability; prefer `net10.0-windows` for controlled modern .NET machines; use multi-targeting when a library must serve both. For NativeAOT preview, prefer a separate or conditionally configured `net10.0-windows`/`win-x64` project until package layout stabilizes.
- Use `ExcelAddInExplicitExports=true` for production examples.
- Include `[ExcelFunction]` and `[ExcelArgument]` metadata.
- Include at least one smoke-testable UDF.
- For async examples, use `Task<T>` or `IObservable<T>` registration patterns and include a cancellation token only when the pattern uses it.
- For ordinary managed ribbon examples, make the ribbon class `public`, derive directly from `ExcelRibbon`, and mark it `[ComVisible(true)]`. For NativeAOT preview ribbon examples, check `references/11_native_aot_preview.md` and prefer the NativeAOT `IExcelRibbon` pattern shown there.

## Source awareness

This first package is based on public Excel-DNA documentation, NuGet metadata, official Microsoft Office/Excel C API documentation, public Excel-DNA repositories/samples, and public clues about skill-package construction. It also flags places where the existing documentation appears stale, incomplete, or sample-only.
