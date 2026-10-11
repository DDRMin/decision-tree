using System;
using DecisionTree.Excel.Diagnostics;
using DecisionTree.Excel.Gateway;
using DecisionTree.Excel.UI;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;

namespace DecisionTree.Excel
{
    public sealed class AddIn : IExcelAddIn
    {
        // Held in a field so the COM event sink is not garbage collected.
        private static Application? _app;

        public void AutoOpen()
        {
            Log.Info("DecisionTree Studio loading (Excel " + ExcelDnaUtil.ExcelVersion + ")");
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try
                {
                    _app = ExcelGateway.App;
                    _app.WorkbookBeforeClose += OnWorkbookBeforeClose;
                    _app.WindowActivate += OnWindowActivate;
                    _app.SheetSelectionChange += OnSheetSelectionChange;
                    PaneManager.Initialize();
                    ExcelComAddInHelper.LoadComAddIn(new ShutdownWatcher());
                    Log.Info("DecisionTree Studio ready");
                }
                catch (Exception ex)
                {
                    Log.Error("Startup failed", ex);
                }
            });
        }

        public void AutoClose() => ReleaseExcel("AutoClose");

        /// <summary>
        /// Drop every reference to Excel objects so EXCEL.EXE can exit. Safe to call more than once.
        /// </summary>
        internal static void ReleaseExcel(string reason)
        {
            try
            {
                PaneManager.Shutdown();
                if (_app != null)
                {
                    _app.WorkbookBeforeClose -= OnWorkbookBeforeClose;
                    _app.WindowActivate -= OnWindowActivate;
                    _app.SheetSelectionChange -= OnSheetSelectionChange;
                    _app = null;
                    Log.Info("Released Excel references (" + reason + ")");
                }
            }
            catch (Exception ex)
            {
                Log.Error("Releasing Excel references failed (" + reason + ")", ex);
            }

            // Release RCWs that are only waiting for finalization.
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }

        private static void OnWorkbookBeforeClose(Workbook wb, ref bool cancel)
        {
            try { PaneManager.OnWorkbookBeforeClose(wb); }
            catch (Exception ex) { Log.Error("WorkbookBeforeClose", ex); }
        }

        private static void OnWindowActivate(Workbook wb, Window window)
        {
            try { Log.Info("Window activated: " + wb.Name + " (hwnd " + window.Hwnd + ")"); }
            catch { }
        }

        private static void OnSheetSelectionChange(object sheet, Range target)
        {
            // Event-based selection, compared against the poll in PaneManager during the spike.
            try { Log.Info("SheetSelectionChange: " + target.Address[false, false]); }
            catch { }
        }
    }
}
