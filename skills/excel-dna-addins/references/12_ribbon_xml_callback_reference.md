# Ribbon XML and Callback Reference

## Contents

- [When to use this reference](#when-to-use-this-reference)
- [Excel-DNA hosting model](#excel-dna-hosting-model)
- [Namespaces and top-level callbacks](#namespaces-and-top-level-callbacks)
- [Control map](#control-map)
- [Callback signature table](#callback-signature-table)
- [Practical callback notes](#practical-callback-notes)
- [Source notes](#source-notes)

## When to use this reference

Open this file when a user asks for Office Ribbon XML, callback signatures, dynamic ribbon state, `IRibbonControl`, `IRibbonUI`, image loading, control attributes, or the details behind an Excel-DNA `ExcelRibbon` implementation.

Excel-DNA loads and exposes the COM add-in that Office calls, but the Ribbon XML schema and callback signatures are Office Custom UI contracts. Treat Excel-DNA as the host/bridge and Microsoft Office Custom UI as the source of truth for control semantics.

## Excel-DNA hosting model

Ordinary managed Excel-DNA ribbon classes should:

- be `public`;
- derive directly from `ExcelDna.Integration.CustomUI.ExcelRibbon`;
- be visible to COM, commonly with `[ComVisible(true)]` or `<ComVisible(True)>`;
- expose public callback methods whose names match the XML callback attributes;
- use `ExcelDnaUtil.Application` only from valid UI, command, or macro contexts, not from thread-safe UDFs or arbitrary background threads.

Minimal C# shape:

```csharp
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using System.Runtime.InteropServices;

[ComVisible(true)]
public class RibbonController : ExcelRibbon
{
    public override string GetCustomUI(string ribbonId)
    {
        return @"
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
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
        dynamic excel = ExcelDnaUtil.Application;
        excel.ActiveCell.Value2 = "Hello from Excel-DNA";
    }
}
```

Excel-DNA's `IRibbonControl` exposes `Id`, `Context`, and `Tag`. `IRibbonUI` exposes `Invalidate`, `InvalidateControl`, `InvalidateControlMso`, and Office 2010 tab activation methods in current Excel-DNA source.

## Namespaces and top-level callbacks

Use the Office 2007 namespace for broad compatibility:

```xml
<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">
```

Use the Office 2010 namespace when you need controls or behavior from the 2010 Custom UI schema:

```xml
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
```

Excel-DNA's `ExcelRibbon` checks the Excel version and chooses 2010 markup first for Excel 2010 or later, with 2007 markup as fallback.

Top-level callbacks:

| Element | Callback | C# | VB.NET | VBA | C++ |
|---|---|---|---|---|---|
| `customUI` | `onLoad` | `void OnLoad(IRibbonUI ribbon)` | `Sub OnLoad(ribbon As IRibbonUI)` | `Sub OnLoad(ribbon As IRibbonUI)` | `HRESULT OnLoad([in] IRibbonUI *pRibbon)` |
| `customUI` | `loadImage` | `IPictureDisp LoadImage(string imageId)` | `Function LoadImage(imageId As String) As IPictureDisp` | `Sub LoadImage(imageId As String, ByRef image)` | `HRESULT LoadImage([in] BSTR *pbstrImageId, [out, retval] IPictureDisp **ppdispImage)` |

Excel-DNA adapts `LoadImage` more flexibly than the raw Office signature: the base class virtual method returns `object`, and comments in the source indicate supported return forms include `IPictureDisp`, `System.Drawing.Bitmap`, or an `imageMso` string.

## Control map

This map is condensed from the Office 2007 Custom UI control reference linked by the Excel-DNA RibbonBasics tutorial. `Common` means standard Office Custom UI attributes and callbacks for that element.

| Element | Purpose | Attributes and callbacks | Children |
|---|---|---|---|
| `customUI` | Root of the Ribbon customization. | none in the 2007 table; top-level callbacks include `onLoad` and `loadImage`. | `commands`, `ribbon` |
| `commands` | Container for globally repurposed built-in commands. | none | `command` |
| `command` | Repurposed built-in command. | `enabled`, `getEnabled`, required `idMso`, `onAction` | none |
| `ribbon` | Main Fluent UI definition. | `startFromScratch` | `contextualTabs`, `officeMenu`, `qat`, `tabs` |
| `contextualTabs` | Context-sensitive tabs, such as picture tools. | none | `tabSet` |
| `tabSet` | Collection of contextual tab controls. | `getVisible`, required `idMso`, `visible` | `tab` |
| `qat` | Quick Access Toolbar. | none | `documentControls`, `sharedControls` |
| `sharedControls` | Shared QAT controls. Prefer `documentControls` for document-specific scenarios. | none | `button`, `control`, `separator` |
| `documentControls` | Document-specific QAT controls. | none | `button`, `control`, `separator` |
| `officeMenu` | Office menu controls. | none | `button`, `checkBox`, `control`, `dynamicMenu`, `gallery`, `menu`, `menuSeparator`, `splitButton`, `toggleButton` |
| `tabs` | Container for normal tabs. | none | `tab` |
| `tab` | Ribbon tab. | `getKeytip`, `getLabel`, `getVisible`, `id`, `idMso`, `idQ`, insertion attributes, `keytip`, `label`, `tag`, `visible` | `group` |
| `group` | Group within a tab. | image, label, screentip, supertip, keytip, visible, id/insertion attributes and corresponding `get*` callbacks | `box`, `button`, `buttonGroup`, `checkBox`, `comboBox`, `control`, `dialogBoxLauncher`, `dropDown`, `editBox`, `gallery`, `labelControl`, `menu`, `separator`, `splitButton`, `toggleButton` |
| `box` | Horizontal or vertical layout container. | `getVisible`, `id`, `idQ`, insertion attributes, `visible` | `box`, `button`, `buttonGroup`, `checkBox`, `comboBox`, `control`, `dropDown`, `dynamicMenu`, `editBox`, `gallery`, `labelControl`, `menu`, `splitButton`, `toggleButton` |
| `button` | Clickable command. | common image/label/help/visibility attributes, `getShowImage`, `getShowLabel`, `getSize`, `onAction` | none |
| `buttonGroup` | Compact grouping of button-like controls. | `getVisible`, `id`, `idQ`, insertion attributes, `visible` | `button`, `control`, `dynamicMenu`, `gallery`, `menu`, `splitButton`, `toggleButton` |
| `checkBox` | Boolean check box. | common label/help/visibility attributes, `getPressed`, `onAction` | none |
| `comboBox` | Editable combo box. | common image/label/help/visibility attributes; shared `editBox` callbacks `getText` and `onChange`; dynamic item callbacks `getItemCount`, `getItemID`, `getItemImage`, `getItemLabel`, `getItemScreentip`, `getItemSupertip`; `maxLength`, `sizeString`, `showItemImage` | `item` |
| `dialogBoxLauncher` | Group dialog launcher. Must be the final element in a group. | none | required `button` |
| `dropDown` | Non-editable drop-down list. | common image/label/help/visibility attributes; dynamic item callbacks shared with `comboBox`; `getSelectedItemID`, `getSelectedItemIndex`, `onAction`, `showItemImage`, `showItemLabel`, `sizeString` | `item` |
| `dynamicMenu` | Runtime-generated menu. | common image/label/help/visibility attributes, `getContent` | menu-like XML returned from `getContent` |
| `editBox` | Text input. | common image/label/help/visibility attributes, `getText`, `maxLength`, `onChange`, `sizeString` | none |
| `gallery` | Gallery of visual/text items. | common image/label/help/visibility attributes; dynamic item callbacks shared with `comboBox`; `columns`, `rows`, `itemHeight`, `itemWidth`, `getItemHeight`, `getItemWidth`, `getSelectedItemID`, `getSelectedItemIndex`, `onAction`, `showItemImage`, `showItemLabel`, `sizeString` | `item`, then optional bottom `button` controls |
| `item` | Static `gallery`, `dropDown`, or `comboBox` item. | `id`, `image`, `imageMso`, `label`, `screentip`, `supertip` | none |
| `labelControl` | Static/dynamic label. | `enabled`, `getEnabled`, `getLabel`, `getScreentip`, `getShowLabel`, `getSupertip`, `getVisible`, id/insertion attributes, `label`, `screentip`, `showLabel`, `supertip`, `tag`, `visible` | none |
| `menu` | Menu container. | common image/label/help/visibility attributes, `getSize` | `button`, `checkBox`, `control`, `dynamicMenu`, `gallery`, `menu`, `menuSeparator`, `splitButton`, `toggleButton` |
| `menuSeparator` | Separator inside a menu, optionally with text. | `id`, `idQ`, insertion attributes, `title`, `getTitle` | none |
| `separator` | Separator between controls. | `getVisible`, `id`, `idQ`, insertion attributes, `visible` | none |
| `splitButton` | Main button plus menu. | `enabled`, `getEnabled`, `getKeytip`, `getShowLabel`, `getSize`, `getSupertip`, `getVisible`, id/insertion attributes, `keytip`, `showLabel`, `size`, `supertip`, `tag`, `visible` | one `button` or `toggleButton`, followed by one `menu` |
| `toggleButton` | Pressed/unpressed button. | common image/label/help/visibility attributes, `getPressed`, `onAction` | none |

The `control` element appears as an allowed child in the Office tables for built-in controls; use `idMso` and the relevant Office Custom UI schema when reusing a built-in Office control.

## Callback signature table

Use these signatures for public methods on the Excel-DNA ribbon class. Callback method names in code may differ from the generic names below; the XML attribute value selects the actual method name.

| Control | Callback | C# | VB.NET | VBA | C++ |
|---|---|---|---|---|---|
| `(several controls)` | `getDescription` | `string GetDescription(IRibbonControl control)` | `Function GetDescription(control As IRibbonControl) As String` | `Sub GetDescription(control As IRibbonControl, ByRef description)` | `HRESULT GetDescription([in] IRibbonControl *pControl, [out, retval] BSTR *pbstrDescription)` |
| `(several controls)` | `getEnabled` | `bool GetEnabled(IRibbonControl control)` | `Function GetEnabled(control As IRibbonControl) As Boolean` | `Sub GetEnabled(control As IRibbonControl, ByRef enabled)` | `HRESULT GetEnabled([in] IRibbonControl *pControl, [out, retval] VARIANT_BOOL *pvarfEnabled)` |
| `(several controls)` | `getImage` | `IPictureDisp GetImage(IRibbonControl control)` | `Function GetImage(control As IRibbonControl) As IPictureDisp` | `Sub GetImage(control As IRibbonControl, ByRef image)` | `HRESULT GetImage([in] IRibbonControl *pControl, [out, retval] IPictureDisp **ppdispImage)` |
| `(several controls)` | `getImageMso` | `string GetImageMso(IRibbonControl control)` | `Function GetImageMso(control As IRibbonControl) As String` | `Sub GetImageMso(control As IRibbonControl, ByRef imageMso)` | `HRESULT GetImageMso([in] IRibbonControl *pControl, [out, retval] BSTR *pbstrImageMso)` |
| `(several controls)` | `getLabel` | `string GetLabel(IRibbonControl control)` | `Function GetLabel(control As IRibbonControl) As String` | `Sub GetLabel(control As IRibbonControl, ByRef label)` | `HRESULT GetLabel([in] IRibbonControl *pControl, [out, retval] BSTR *pbstrLabel)` |
| `(several controls)` | `getKeytip` | `string GetKeytip(IRibbonControl control)` | `Function GetKeytip(control As IRibbonControl) As String` | `Sub GetKeytip(control As IRibbonControl, ByRef label)` | `HRESULT GetKeytip([in] IRibbonControl *pControl, [out, retval] BSTR *pbstrKeytip)` |
| `(several controls)` | `getSize` | `RibbonControlSize GetSize(IRibbonControl control)` | `Function GetSize(control As IRibbonControl) As RibbonControlSize` | `Sub GetSize(control As IRibbonControl, ByRef size)` | `HRESULT GetSize([in] IRibbonControl *pControl, [out, retval] RibbonControlSize *pintSize)` |
| `(several controls)` | `getScreentip` | `string GetScreentip(IRibbonControl control)` | `Function GetScreentip(control As IRibbonControl) As String` | `Sub GetScreentip(control As IRibbonControl, ByRef screentip)` | `HRESULT GetScreentip([in] IRibbonControl *pControl, [out, retval] BSTR *pbstrScreentip)` |
| `(several controls)` | `getSupertip` | `string GetSupertip(IRibbonControl control)` | `Function GetSupertip(control As IRibbonControl) As String` | `Sub GetSupertip(control As IRibbonControl, ByRef screentip)` | `HRESULT GetSupertip([in] IRibbonControl *pControl, [out, retval] BSTR *pbstrScreentip)` |
| `(several controls)` | `getVisible` | `bool GetVisible(IRibbonControl control)` | `Function GetVisible(control As IRibbonControl) As Boolean` | `Sub GetVisible(control As IRibbonControl, ByRef visible)` | `HRESULT GetVisible([in] IRibbonControl *pControl, [out, retval] VARIANT_BOOL *pvarfVisible)` |
| `button` | `getShowImage` | `bool GetShowImage(IRibbonControl control)` | `Function GetShowImage(control As IRibbonControl) As Boolean` | `Sub GetShowImage(control As IRibbonControl, ByRef showImage)` | `HRESULT GetShowImage([in] IRibbonControl *pControl, [out, retval] VARIANT_BOOL *pvarShowImage)` |
| `button` | `getShowLabel` | `bool GetShowLabel(IRibbonControl control)` | `Function GetShowLabel(control As IRibbonControl) As Boolean` | `Sub GetShowLabel(control As IRibbonControl, ByRef showLabel)` | `HRESULT GetShowLabel([in] IRibbonControl *pControl, [out, retval] VARIANT_BOOL *pvarShowLabel)` |
| `button` | `onAction` | `void OnAction(IRibbonControl control)` | `Sub OnAction(control As IRibbonControl)` | `Sub OnAction(control As IRibbonControl)` | `HRESULT OnAction([in] IRibbonControl *pControl)` |
| `button` | `onAction` repurposed | `void OnAction(IRibbonControl control, ref bool cancelDefault)` | `Sub OnAction(control As IRibbonControl, ByRef cancelDefault)` | `Sub OnAction(control As IRibbonControl, ByRef cancelDefault)` | `HRESULT OnAction([in] IRibbonControl *pControl, [in,out] VARIANT_BOOL *fCancelDefault)` |
| `checkBox` | `getPressed` | `bool GetPressed(IRibbonControl control)` | `Function GetPressed(control As IRibbonControl) As Boolean` | `Sub GetPressed(control As IRibbonControl, ByRef returnValue)` | `HRESULT GetPressed([in] IRibbonControl *pControl, [out, retval] VARIANT_BOOL *pvarfPressed)` |
| `checkBox` | `onAction` | `void OnAction(IRibbonControl control, bool pressed)` | `Sub OnAction(control As IRibbonControl, pressed As Boolean)` | `Sub OnAction(control As IRibbonControl, pressed As Boolean)` | `HRESULT OnAction([in] IRibbonControl *pControl, [in] VARIANT_BOOL *pvarfPressed)` |
| `comboBox` | `getItemCount` | `int GetItemCount(IRibbonControl control)` | `Function GetItemCount(control As IRibbonControl) As Integer` | `Sub GetItemCount(control As IRibbonControl, ByRef count)` | `HRESULT GetItemCount([in] IRibbonControl *pControl, [out, retval] long *count)` |
| `comboBox` | `getItemID` | `string GetItemID(IRibbonControl control, int index)` | `Function GetItemID(control As IRibbonControl, index As Integer) As String` | `Sub GetItemID(control As IRibbonControl, index As Integer, ByRef id)` | `HRESULT GetItemID([in] IRibbonControl *pControl, [in] long cIndex, [out, retval] BSTR *pbstrID)` |
| `comboBox` | `getItemImage` | `IPictureDisp GetItemImage(IRibbonControl control, int index)` | `Function GetItemImage(control As IRibbonControl, index As Integer) As IPictureDisp` | `Sub GetItemImage(control As IRibbonControl, index As Integer, ByRef image)` | `HRESULT GetItemImage([in] IRibbonControl *pControl, [in] long cIndex, [out, retval] IPictureDisp **ppdispImage)` |
| `comboBox` | `getItemLabel` | `string GetItemLabel(IRibbonControl control, int index)` | `Function GetItemLabel(control As IRibbonControl, index As Integer) As String` | `Sub GetItemLabel(control As IRibbonControl, index As Integer, ByRef label)` | `HRESULT GetItemLabel([in] IRibbonControl *pControl, [in] long cIndex, [out, retval] BSTR *pbstrLabel)` |
| `comboBox` | `getItemScreenTip` | `string GetItemScreenTip(IRibbonControl control, int index)` | `Function GetItemScreentip(control As IRibbonControl, index As Integer) As String` | `Sub GetItemScreenTip(control As IRibbonControl, index As Integer, ByRef screentip)` | `HRESULT GetItemScreentip([in] IRibbonControl *pControl, [in] long cIndex, [out, retval] BSTR *pbstrScreentip)` |
| `comboBox` | `getItemSuperTip` | `string GetItemSuperTip(IRibbonControl control, int index)` | `Function GetItemSuperTip(control As IRibbonControl, index As Integer) As String` | `Sub GetItemSuperTip(control As IRibbonControl, index As Integer, ByRef supertip)` | `HRESULT GetItemSuperTip([in] IRibbonControl *pControl, [in] long cIndex, [out, retval] BSTR *pbstrSupertip)` |
| `comboBox` | `getText` | `string GetText(IRibbonControl control)` | `Function GetText(control As IRibbonControl) As String` | `Sub GetText(control As IRibbonControl, ByRef text)` | `HRESULT GetText([in] IRibbonControl *pControl, [out, retval] BSTR *pbstrText)` |
| `comboBox` | `onChange` | `void OnChange(IRibbonControl control, string text)` | `Sub OnChange(control As IRibbonControl, text As String)` | `Sub OnChange(control As IRibbonControl, text As String)` | `HRESULT OnChange([in] IRibbonControl *pControl, [in] BSTR *pbstrText)` |
| `dropDown` | `getItemCount` | `int GetItemCount(IRibbonControl control)` | `Function GetItemCount(control As IRibbonControl) As Integer` | `Sub GetItemCount(control As IRibbonControl, ByRef count)` | `HRESULT GetItemCount([in] IRibbonControl *pControl, [out, retval] long *count)` |
| `dropDown` | `getItemID` | `string GetItemID(IRibbonControl control, int index)` | `Function GetItemID(control As IRibbonControl, index As Integer) As String` | `Sub GetItemID(control As IRibbonControl, index As Integer, ByRef id)` | `HRESULT GetItemID([in] IRibbonControl *pControl, [in] long cIndex, [out, retval] BSTR *pbstrID)` |
| `dropDown` | `getItemImage` | `IPictureDisp GetItemImage(IRibbonControl control, int index)` | `Function GetItemImage(control As IRibbonControl, index As Integer) As IPictureDisp` | `Sub GetItemImage(control As IRibbonControl, index As Integer, ByRef image)` | `HRESULT GetItemImage([in] IRibbonControl *pControl, [in] long cIndex, [out, retval] IPictureDisp **ppdispImage)` |
| `dropDown` | `getItemLabel` | `string GetItemLabel(IRibbonControl control, int index)` | `Function GetItemLabel(control As IRibbonControl, index As Integer) As String` | `Sub GetItemLabel(control As IRibbonControl, index As Integer, ByRef label)` | `HRESULT GetItemLabel([in] IRibbonControl *pControl, [in] long cIndex, [out, retval] BSTR *pbstrLabel)` |
| `dropDown` | `getItemScreenTip` | `string GetItemScreenTip(IRibbonControl control, int index)` | `Function GetItemScreentip(control As IRibbonControl, index As Integer) As String` | `Sub GetItemScreenTip(control As IRibbonControl, index As Integer, ByRef screenTip)` | `HRESULT GetItemScreentip([in] IRibbonControl *pControl, [in] long cIndex, [out, retval] BSTR *pbstrScreentip)` |
| `dropDown` | `getItemSuperTip` | `string GetItemSuperTip(IRibbonControl control, int index)` | `Function GetItemSuperTip(control As IRibbonControl, index As Integer) As String` | `Sub GetItemSuperTip(control As IRibbonControl, index As Integer, ByRef superTip)` | `HRESULT GetItemSuperTip([in] IRibbonControl *pControl, [in] long cIndex, [out, retval] BSTR *pbstrSupertip)` |
| `dropDown` | `getSelectedItemID` | `string GetSelectedItemID(IRibbonControl control)` | `Function GetSelectedItemID(control As IRibbonControl) As String` | `Sub GetSelectedItemID(control As IRibbonControl, ByRef selectedId)` | `HRESULT GetSelectedItemID([in] IRibbonControl *pControl, [out, retval] BSTR *pbstrSelectedId)` |
| `dropDown` | `getSelectedItemIndex` | `int GetSelectedItemIndex(IRibbonControl control)` | `Function GetSelectedItemIndex(control As IRibbonControl) As Integer` | `Sub GetSelectedItemIndex(control As IRibbonControl, ByRef index)` | `HRESULT GetSelectedItemIndex([in] IRibbonControl *pControl, [out, retval] LONG *pcItemIndex)` |
| `dropDown` | `onAction` | `void OnAction(IRibbonControl control, string selectedId, int selectedIndex)` | `Sub OnAction(control As IRibbonControl, selectedId As String, selectedIndex As Integer)` | `Sub OnAction(control As IRibbonControl, selectedId As String, selectedIndex As Integer)` | `HRESULT OnAction([in] IRibbonControl *pControl, [in] BSTR *selectedId, [in] LONG cSelectedIndex)` |
| `dynamicMenu` | `getContent` | `string GetContent(IRibbonControl control)` | `Function GetContent(control As IRibbonControl) As String` | `Sub GetContent(control As IRibbonControl, ByRef content)` | `HRESULT GetContent([in] IRibbonControl *pControl, [out, retval] BSTR *pbstrContent)` |
| `editBox` | `getText` | `string GetText(IRibbonControl control)` | `Function GetText(control As IRibbonControl) As String` | `Sub GetText(control As IRibbonControl, ByRef text)` | `HRESULT GetText([in] IRibbonControl *pControl, [out, retval] BSTR *pbstrText)` |
| `editBox` | `onChange` | `void OnChange(IRibbonControl control, string text)` | `Sub OnChange(control As IRibbonControl, text As String)` | `Sub OnChange(control As IRibbonControl, text As String)` | `HRESULT OnChange([in] IRibbonControl *pControl, [in] BSTR *pbstrText)` |
| `gallery` | `getItemCount` | `int GetItemCount(IRibbonControl control)` | `Function GetItemCount(control As IRibbonControl) As Integer` | `Sub GetItemCount(control As IRibbonControl, ByRef count)` | `HRESULT GetItemCount([in] IRibbonControl *pControl, [out, retval] long *count)` |
| `gallery` | `getItemHeight` | `int GetItemHeight(IRibbonControl control)` | `Function GetItemHeight(control As IRibbonControl) As Integer` | `Sub GetItemHeight(control As IRibbonControl, ByRef height)` | `HRESULT GetItemHeight([in] IRibbonControl *pControl, [out, retval] LONG *height)` |
| `gallery` | `getItemID` | `string GetItemID(IRibbonControl control, int index)` | `Function GetItemID(control As IRibbonControl, index As Integer) As String` | `Sub GetItemID(control As IRibbonControl, index As Integer, ByRef id)` | `HRESULT GetItemID([in] IRibbonControl *pControl, [in] long cIndex, [out, retval] BSTR *pbstrID)` |
| `gallery` | `getItemImage` | `IPictureDisp GetItemImage(IRibbonControl control, int index)` | `Function GetItemImage(control As IRibbonControl, index As Integer) As IPictureDisp` | `Sub GetItemImage(control As IRibbonControl, index As Integer, ByRef image)` | `HRESULT GetItemImage([in] IRibbonControl *pControl, [in] long cIndex, [out, retval] IPictureDisp **ppdispImage)` |
| `gallery` | `getItemLabel` | `string GetItemLabel(IRibbonControl control, int index)` | `Function GetItemLabel(control As IRibbonControl, index As Integer) As String` | `Sub GetItemLabel(control As IRibbonControl, index As Integer, ByRef label)` | `HRESULT GetItemLabel([in] IRibbonControl *pControl, [in] long cIndex, [out, retval] BSTR *pbstrLabel)` |
| `gallery` | `getItemScreenTip` | `string GetItemScreenTip(IRibbonControl control, int index)` | `Function GetItemScreentip(control As IRibbonControl, index As Integer) As String` | `Sub GetItemScreenTip(control As IRibbonControl, index As Integer, ByRef screenTip)` | `HRESULT GetItemScreentip([in] IRibbonControl *pControl, [in] long cIndex, [out, retval] BSTR *pbstrScreentip)` |
| `gallery` | `getItemSuperTip` | `string GetItemSuperTip(IRibbonControl control, int index)` | `Function GetItemSuperTip(control As IRibbonControl, index As Integer) As String` | `Sub GetItemSuperTip(control As IRibbonControl, index As Integer, ByRef superTip)` | `HRESULT GetItemSuperTip([in] IRibbonControl *pControl, [in] long cIndex, [out, retval] BSTR *pbstrSupertip)` |
| `gallery` | `getItemWidth` | `int GetItemWidth(IRibbonControl control)` | `Function GetItemWidth(control As IRibbonControl) As Integer` | `Sub GetItemWidth(control As IRibbonControl, ByRef width)` | `HRESULT GetItemWidth([in] IRibbonControl *pControl, [out, retval] LONG *width)` |
| `gallery` | `getSelectedItemID` | `string GetSelectedItemID(IRibbonControl control)` | `Function GetSelectedItemID(control As IRibbonControl) As String` | `Sub GetSelectedItemID(control As IRibbonControl, ByRef selectedId)` | `HRESULT GetSelectedItemID([in] IRibbonControl *pControl, [out, retval] BSTR *pbstrSelectedId)` |
| `gallery` | `getSelectedItemIndex` | `int GetSelectedItemIndex(IRibbonControl control)` | `Function GetSelectedItemIndex(control As IRibbonControl) As Integer` | `Sub GetSelectedItemIndex(control As IRibbonControl, ByRef index)` | `HRESULT GetSelectedItemIndex([in] IRibbonControl *pControl, [out, retval] LONG *pcItemIndex)` |
| `gallery` | `onAction` | `void OnAction(IRibbonControl control, string selectedId, int selectedIndex)` | `Sub OnAction(control As IRibbonControl, selectedId As String, selectedIndex As Integer)` | `Sub OnAction(control As IRibbonControl, selectedId As String, selectedIndex As Integer)` | `HRESULT OnAction([in] IRibbonControl *pControl, [in] BSTR *selectedId, [in] LONG cSelectedIndex)` |
| `menuSeparator` | `getTitle` | `string GetTitle(IRibbonControl control)` | `Function GetTitle(control As IRibbonControl) As String` | `Sub GetTitle(control As IRibbonControl, ByRef title)` | `HRESULT GetTitle([in] IRibbonControl *pControl, [out, retval] BSTR *pbstrTitle)` |
| `toggleButton` | `getPressed` | `bool GetPressed(IRibbonControl control)` | `Function GetPressed(control As IRibbonControl) As Boolean` | `Sub GetPressed(control As IRibbonControl, ByRef returnValue)` | `HRESULT GetPressed([in] IRibbonControl *pControl, [out, retval] VARIANT_BOOL *pvarfPressed)` |
| `toggleButton` | `onAction` | `void OnAction(IRibbonControl control, bool pressed)` | `Sub OnAction(control As IRibbonControl, pressed As Boolean)` | `Sub OnAction(control As IRibbonControl, pressed As Boolean)` | `HRESULT OnAction([in] IRibbonControl *pControl, [in] VARIANT_BOOL *pvarfPressed)` |
| `toggleButton` | `onAction` repurposed | `void OnAction(IRibbonControl control, bool pressed, ref bool cancelDefault)` | `Sub OnAction(control As IRibbonControl, pressed As Boolean, ByRef cancelDefault)` | `Sub OnAction(control As IRibbonControl, pressed As Boolean, ByRef cancelDefault)` | `HRESULT OnAction([in] IRibbonControl *pControl, [in] VARIANT_BOOL *pvarfPressed, [in,out] VARIANT_BOOL *fCancelDefault)` |

The Microsoft Part 3 table has a few visible transcription inconsistencies around `getSelectedItemID`, where some non-C# columns look like `index`/integer returns. Prefer the semantically consistent form above: `getSelectedItemID` returns the selected item ID string, while `getSelectedItemIndex` returns the selected item index.

## Practical callback notes

- Keep callback names unique within the add-in. Do not rely on overloads with the same name and different signatures; the Office COM callback dispatch path can pick the wrong overload.
- Use `control.Id` for ordinary control routing and `control.Tag` when you need an arbitrary string that is not constrained like an XML ID.
- Dynamic state callbacks such as `getEnabled`, `getVisible`, `getLabel`, and gallery item callbacks are cached by Office. Store the `IRibbonUI` from `onLoad` and call `Invalidate()` or `InvalidateControl(id)` when state changes.
- `loadImage` is called for image names supplied by XML `image` attributes; it is not the same as a per-control dynamic `getImage` callback.
- `dynamicMenu getContent` must return valid Custom UI XML for menu content. Keep the returned markup small and deterministic.
- For Excel-DNA, execute workbook mutations from ribbon callbacks, commands, or queued macros. Do not push ribbon-triggered workbook mutation back into worksheet UDF calculation.

## Source notes

Primary sources used:

- Excel-DNA Tutorials `Fundamentals/RibbonBasics` README and resource example.
- Microsoft previous-version Office documentation linked by the tutorial: Part 1 overview, Part 2 controls and callbacks, and Part 3 FAQ/signature table.
- Current Excel-DNA source for `ExcelRibbon`, `IRibbonControl`, `IRibbonUI`, and NativeAOT `IExcelRibbon` shape.
