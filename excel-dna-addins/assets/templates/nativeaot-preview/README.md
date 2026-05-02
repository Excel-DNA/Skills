# NativeAOT Preview Excel-DNA Template

This is a minimal NativeAOT preview add-in template.

## Build and publish

```bash
dotnet restore -r win-x64
dotnet publish -c Release -r win-x64
```

Load the published `*-AddIn64.xll` from the publish folder in 64-bit Excel.

## Test formulas

```excel
=NATIVE.ADD(1,2)
=NATIVE.HELLO("Excel")
```

## Review before using in production

- NativeAOT support is a preview/specialized path.
- The template references `ExcelDna.AddIn.NativeAOT`, not ordinary `ExcelDna.AddIn`.
- Treat all AOT and trimming warnings as defects until investigated.
- Do not add `ExcelDna.IntelliSense`, custom task panes, ribbon images, plugin loading, or reflection-heavy dependencies without a focused compatibility test.
- Test the published `*-AddIn64.xll` on a clean 64-bit Excel machine.
