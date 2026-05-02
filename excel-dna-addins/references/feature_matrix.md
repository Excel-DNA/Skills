# Feature Matrix

| Feature | Excel-DNA support pattern | Primary reference |
|---|---|---|
| Scalar UDFs | `[ExcelFunction] public static` methods | `03_udf_core_patterns.md` |
| Array/range UDFs | `object[,]`, `ExcelReference` | `03_udf_core_patterns.md`, `07_com_and_c_api.md` |
| Optional parameters | defaults or `ExcelMissing`/registration support | `04_function_registration_extended.md` |
| Async one-shot | `Task<T>` or `ExcelAsyncUtil.Run` | `05_async_streaming_rtd.md` |
| Streaming | `IObservable<T>` / RTD | `05_async_streaming_rtd.md` |
| Object handles | handle registration/RTD patterns | `04_function_registration_extended.md` |
| Ribbon | `ExcelRibbon` subclass and CustomUI XML | `06_ribbon_ctp_intellisense.md` |
| Custom task panes | CTP manager + UI control | `06_ribbon_ctp_intellisense.md` |
| IntelliSense | ExcelDna.IntelliSense + metadata | `06_ribbon_ctp_intellisense.md` |
| COM automation | `ExcelDnaUtil.Application` in safe contexts | `07_com_and_c_api.md` |
| C API calls | `XlCall`, `ExcelReference` | `07_com_and_c_api.md` |
| Testing | ExcelDna.Testing, xUnit, smoke workbooks | `08_testing_distribution_installation.md` |
| Distribution | packed `.xll`, signing, WiX | `08_testing_distribution_installation.md` |
| NativeAOT runtime-free 64-bit XLL | `ExcelDna.AddIn.NativeAOT` preview, `net10.0-windows`, `win-x64`, `PublishAot` | `11_native_aot_preview.md` |
| NativeAOT UDFs | explicitly attributed static methods; publish and load `*-AddIn64.xll` | `11_native_aot_preview.md`, `03_udf_core_patterns.md` |
| NativeAOT async/functions/object handles | preview-supported patterns requiring AOT-warning-clean publish tests | `11_native_aot_preview.md`, `05_async_streaming_rtd.md`, `04_function_registration_extended.md` |
| NativeAOT UI limitations | simple ribbon via `IExcelRibbon`; IntelliSense overlay, CTPs, ribbon images treated as unsupported/unproven | `11_native_aot_preview.md`, `06_ribbon_ctp_intellisense.md` |
| Troubleshooting | load/runtime/trust/registration matrix | `09_troubleshooting_and_known_issues.md` |
