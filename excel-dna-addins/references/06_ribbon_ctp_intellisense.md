# Ribbon, Custom Task Panes, IntelliSense, and COM-visible UI Surfaces

## Contents

- [Ribbon support](#ribbon-support)
- [Ribbon debugging](#ribbon-debugging)
- [Accessing Excel object model from UI callbacks](#accessing-excel-object-model-from-ui-callbacks)
- [Custom task panes](#custom-task-panes)
- [Function IntelliSense](#function-intellisense)
- [UI plus UDF design](#ui-plus-udf-design)
- [NativeAOT preview UI guidance](#nativeaot-preview-ui-guidance)

## Ribbon support

Excel-DNA supports custom ribbon UI through a public class deriving directly from `ExcelDna.Integration.CustomUI.ExcelRibbon`. The class is registered/loaded as a COM add-in by Excel-DNA during add-in startup.

Minimal pattern:

```csharp
using ExcelDna.Integration.CustomUI;
using System.Runtime.InteropServices;

[ComVisible(true)]
public class RibbonController : ExcelRibbon
{
    public override string GetCustomUI(string ribbonId)
    {
        return @"
<customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui'>
  <ribbon>
    <tabs>
      <tab id='MyTab' label='My Add-in'>
        <group id='MyGroup' label='Tools'>
          <button id='HelloButton' label='Say Hello' onAction='OnHello' />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
    }

    public void OnHello(IRibbonControl control)
    {
        System.Windows.Forms.MessageBox.Show("Hello from Excel-DNA");
    }
}
```

Rules:

- Ribbon class must be `public`.
- Derive directly from `ExcelRibbon`.
- Mark COM visible where needed.
- Use the correct Office CustomUI XML namespace.
- Keep callback names stable.
- Use ribbon callbacks for user actions, not worksheet UDFs.

## Ribbon debugging

Common causes of invisible ribbon:

- Add-in not loaded.
- COM add-in helper disabled by Excel.
- Ribbon class not public or not directly deriving from `ExcelRibbon`.
- XML namespace/version mismatch.
- XML syntax error.
- Callback signature mismatch.
- Office UI errors are hidden unless enabled in Excel options.

Troubleshooting steps:

1. Confirm a simple UDF from the same add-in works.
2. Enable Office UI error display.
3. Validate CustomUI XML.
4. Confirm class visibility and `ComVisible`.
5. Check disabled COM add-ins in Excel.

## Accessing Excel object model from UI callbacks

Ribbon callbacks occur on the Excel UI thread and can use the COM object model through `ExcelDnaUtil.Application`.

```csharp
public void SelectA1(IRibbonControl control)
{
    dynamic excel = ExcelDnaUtil.Application;
    excel.ActiveSheet.Range("A1").Select();
}
```

Do not use this pattern from background threads or ordinary thread-safe UDFs.

## Custom task panes

Custom task panes are useful for:

- configuration panels
- function browsers
- model diagnostics
- log/status views
- authentication flows
- workbook-specific navigation

A typical CTP architecture:

- Ribbon button toggles the pane.
- CTP manager creates, tracks, and disposes panes.
- Pane content is a Windows Forms or WPF-hosted control as supported by the chosen runtime and Excel-DNA surface.
- Pane code interacts with Excel only on the main thread.
- Workbook-specific state is keyed by workbook/window identity.

Design concerns:

- Avoid leaking panes when workbooks close.
- Define whether pane state is application-wide or workbook-specific.
- Do not let pane actions mutate workbooks during calculation.
- Keep UI and calculation services separate.

## Function IntelliSense

Excel-DNA IntelliSense provides in-sheet function help for UDFs by overlaying UI because Excel does not provide a general public interface for UDF IntelliSense equivalent to native functions.

Use it to improve discoverability of:

- function names
- argument names
- descriptions
- argument help

Good IntelliSense depends on good `[ExcelFunction]` and `[ExcelArgument]` metadata. Treat metadata as product content, not decoration.

Checklist:

- All public UDFs have names and descriptions.
- Arguments have names and descriptions.
- Categories are consistent.
- Optional/default arguments are documented.
- Examples are available in docs/sample workbook.

## UI plus UDF design

A good Excel-DNA add-in uses UI to support formulas:

- Ribbon inserts example formulas or opens help.
- Task pane lists available functions and diagnostics.
- IntelliSense guides entry.
- UDFs remain the calculation interface.

Avoid making the ribbon the only way to run a model if formulas would make dependencies clearer.


## NativeAOT preview UI guidance

NativeAOT preview support is UDF-first. Treat UI surfaces as feature-by-feature experiments until proven.

Ribbon:

- Ordinary managed Excel-DNA ribbon guidance uses a public class deriving from `ExcelRibbon`.
- NativeAOT preview docs show an `IExcelRibbon` pattern. Use the NativeAOT-specific pattern for preview examples.
- Avoid ribbon images in NativeAOT preview because current docs list loading images for ribbon controls as unsupported.

Excel object model access:

- Prefer `ExcelDnaUtil.DynamicApplication` and `IDynamic` in NativeAOT preview code.
- Do not assume ordinary COM interop, PIAs, or C# `dynamic` behave like framework-dependent .NET.

Custom task panes and rich UI:

- Treat CTP, WinForms, and WPF support as unproven for NativeAOT preview unless the current source/tests specifically confirm the stack.
- Do not recommend NativeAOT for an add-in whose primary value is a custom task pane or rich Windows UI.

Function IntelliSense:

- Do not include `ExcelDna.IntelliSense` in NativeAOT preview starter projects.
- Use `[ExcelFunction]` and `[ExcelArgument]` metadata so the Function Wizard and formula metadata remain useful.
- Public preview discussion indicates ExcelDna.IntelliSense overlay support depends on additional WinForms/trim compatibility work.
