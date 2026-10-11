using System;
using System.Runtime.InteropServices;
using ExcelDna.Integration.Extensibility;
using ExcelDna.Integration;

namespace DecisionTree.Excel
{
    /// <summary>
    /// Excel-DNA's AutoClose is not called when Excel itself shuts down, so references the
    /// add-in holds (event sinks, task panes, cached RCWs) would keep EXCEL.EXE alive.
    /// This COM add-in gets OnBeginShutdown/OnDisconnection from Excel and releases them.
    /// </summary>
    [ComVisible(true)]
    public class ShutdownWatcher : ExcelComAddIn
    {
        public override void OnBeginShutdown(ref Array custom) => AddIn.ReleaseExcel("OnBeginShutdown");

        public override void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom) =>
            AddIn.ReleaseExcel("OnDisconnection");
    }
}
