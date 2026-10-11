using System;
using System.Runtime.InteropServices;
using DecisionTree.Excel.Diagnostics;
using DecisionTree.Excel.UI;
using ExcelDna.Integration.CustomUI;

namespace DecisionTree.Excel
{
    [ComVisible(true)]
    public class Ribbon : ExcelRibbon
    {
        public override string GetCustomUI(string ribbonId) => @"
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
  <ribbon>
    <tabs>
      <tab id='dtsTab' label='DecisionTree Studio'>
        <group id='dtsPanelGroup' label='Panel'>
          <button id='dtsShowPanel' label='Show Panel' size='large' imageMso='ViewsLayoutView'
                  screentip='Show or hide the DecisionTree Studio panel for this window'
                  onAction='OnShowPanel' />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";

        public void OnShowPanel(IRibbonControl control)
        {
            try { PaneManager.ToggleForActiveWindow(); }
            catch (Exception ex) { Log.Error("Show Panel failed", ex); }
        }
    }
}
