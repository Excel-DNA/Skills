# Common Task Recipes

## Create a new UDF add-in

1. Choose target framework using `references/01_decision_guide.md`.
2. Scaffold project.
3. Add `[ExcelFunction]` methods.
4. Build and load `.xll`.
5. Add smoke workbook.
6. Add test project.
7. Decide packing/signing/install path.

## Add async function to existing add-in

1. Confirm Excel-DNA version supports modern registration.
2. Use `Task<T>` for one-shot async; `IObservable<T>` for streaming.
3. Include all inputs in function identity/signature.
4. Add cancellation/timeout.
5. Map service errors deliberately.
6. Test recalculation, deletion, and workbook close.

## Add ribbon command

1. Add public `[ComVisible(true)]` class deriving from `ExcelRibbon`.
2. Return valid CustomUI XML.
3. Add callback methods.
4. Keep workbook mutation in callback/macro context.
5. Enable UI errors while debugging.
6. Test load failure/disabled COM add-in path.

## Diagnose add-in load failure

1. Bitness.
2. Trust/MOTW.
3. Disabled add-ins.
4. Runtime/dependencies.
5. Pack state.
6. Clean machine reproduction.
7. Logs and minimal function sample.

## Review a production Excel-DNA project

Check:

- target frameworks and runtime lifecycle
- explicit exports
- UDF metadata completeness
- thread-safety claims
- async/streaming lifecycle
- COM usage contexts
- object handle lifecycle
- tests
- packing/signing
- installer/runtime prerequisites
- docs/sample workbook


## Create a NativeAOT preview add-in

1. Confirm the requirement is 64-bit Excel for Windows and runtime-free distribution.
2. Open `references/11_native_aot_preview.md` before generating code.
3. Use a NativeAOT-only project with `ExcelDna.AddIn.NativeAOT`, `net10.0-windows`, `win-x64`, and `PublishAot`.
4. Enable explicit exports and write one small attributed smoke-test UDF.
5. Run `dotnet publish -c Release -r win-x64`; treat AOT/trimming warnings as defects.
6. Load the published `*-AddIn64.xll` in 64-bit Excel, preferably on a clean machine without the Desktop Runtime.
7. Sign and package the published artifact only after smoke tests pass.

## Migrate an existing add-in to NativeAOT preview

1. Inventory UDFs, registration transforms, async/streaming functions, object handles, ribbons, CTPs, IntelliSense, COM use, and dependencies.
2. Split a small NativeAOT pilot project rather than changing the production add-in first.
3. Remove ordinary `ExcelDna.AddIn` from the NativeAOT project unless references are conditional and version-aligned.
4. Replace reflection-heavy, dynamic, plugin, serializer, proxy, and expression-compilation paths with static/source-generated alternatives.
5. Use NativeAOT preview UI patterns, especially `IExcelRibbon` and `ExcelDnaUtil.DynamicApplication`/`IDynamic`.
6. Publish frequently and clear every AOT/trimming warning before expanding scope.
7. Keep a managed fallback build until the preview path has passed clean-machine Excel tests.

## Diagnose a NativeAOT preview failure

1. Check that the project references `ExcelDna.AddIn.NativeAOT` and not an unconditional mix of ordinary and NativeAOT packages.
2. Check `TargetFramework`, `RuntimeIdentifier`, and `PublishAot`.
3. Confirm Excel is 64-bit and the file loaded is the published `*-AddIn64.xll`.
4. Review publish warnings such as `IL2026`, `IL3050`, and `IL3058`.
5. Remove `ExcelDna.IntelliSense`, CTPs, ribbon images, dynamic assembly loading, and rich COM assumptions until a minimal repro works.
6. If loader/tool path errors mention missing base XLL assets, check the preview `ExcelDnaToolsPath` workaround and prefer a NativeAOT-only project.
7. Reduce to one UDF, then add async, object handles, ribbon, and dependencies one at a time.
