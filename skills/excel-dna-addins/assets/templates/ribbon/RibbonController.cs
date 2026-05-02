using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using System.Runtime.InteropServices;

namespace MyCompany.MyAddIn;

[ComVisible(true)]
public class RibbonController : ExcelRibbon
{
    public override string GetCustomUI(string ribbonId)
    {
        return @"
<customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui'>
  <ribbon>
    <tabs>
      <tab id='MyAddInTab' label='My Add-in'>
        <group id='ToolsGroup' label='Tools'>
          <button id='InsertFormulaButton' label='Insert Formula' onAction='OnInsertFormula' />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
    }

    public void OnInsertFormula(IRibbonControl control)
    {
        dynamic excel = ExcelDnaUtil.Application;
        excel.ActiveCell.Formula = "=MY.ADD(1,2)";
    }
}
