using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using DecisionTree.Excel.Diagnostics;
using DecisionTree.Excel.Gateway;
using ExcelDna.Integration;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;

namespace DecisionTree.Excel.Spike
{
    /// <summary>
    /// Phase 0 checks against the real Excel host. Each one logs its result.
    /// Throw-away code: the production equivalents live in Persistence/ and Rendering/.
    /// </summary>
    public static class SpikeChecks
    {
        private const string XmlNamespace = "urn:decisiontreestudio:spike:v1";
        private const string ShapePrefix = "DTS_spike_";
        private const string SheetName = "DTS Spike";

        // ---- Check 3: custom XML part survives save / close / reopen ----

        public static void WriteXmlPart() => ExcelGateway.Post("Write XML part", app =>
        {
            var wb = RequireWorkbook(app);
            if (wb == null) return;

            foreach (CustomXMLPart old in wb.CustomXMLParts.SelectByNamespace(XmlNamespace))
                old.Delete();

            var stamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            wb.CustomXMLParts.Add("<spike xmlns=\"" + XmlNamespace + "\"><written>" + stamp + "</written></spike>");
            Log.Info("XML part written to '" + wb.Name + "' at " + stamp + ". Save as .xlsx, close, reopen, then Read.");
        });

        public static void ReadXmlPart() => ExcelGateway.Post("Read XML part", app =>
        {
            var wb = RequireWorkbook(app);
            if (wb == null) return;

            var parts = wb.CustomXMLParts.SelectByNamespace(XmlNamespace);
            if (parts.Count == 0)
            {
                Log.Warn("No spike XML part in '" + wb.Name + "'.");
                return;
            }

            foreach (CustomXMLPart part in parts)
                Log.Info("Found in '" + wb.Name + "': " + part.XML);
        });

        // ---- Check 6: rendering speed (3 shapes per node) ----

        public static void Render(int nodes) => ExcelGateway.Post("Render " + nodes + " nodes", app =>
        {
            var wb = RequireWorkbook(app);
            if (wb == null) return;

            var sheet = FindOrAddSheet(wb);
            sheet.Activate();
            DeleteOwnedShapes(sheet);

            var calculation = app.Calculation;
            app.Calculation = XlCalculation.xlCalculationManual;
            var watch = Stopwatch.StartNew();
            try
            {
                const float cell = 15f;
                var shapes = sheet.Shapes;
                var labels = new object[nodes, 1];
                for (var i = 0; i < nodes; i++)
                {
                    var depth = i % 8;
                    var x = 20f + depth * 4 * cell;
                    var y = 20f + i * cell;

                    var node = shapes.AddShape(depth % 2 == 0 ? MsoAutoShapeType.msoShapeRectangle : MsoAutoShapeType.msoShapeOval,
                        x, y, cell - 2, cell - 2);
                    node.Name = ShapePrefix + "N" + i;
                    node.AlternativeText = "dts:spike";
                    Marshal.ReleaseComObject(node);

                    var branch = shapes.AddLine(x - 3 * cell, y + cell / 2, x, y + cell / 2);
                    branch.Name = ShapePrefix + "B" + i;
                    Marshal.ReleaseComObject(branch);

                    var connector = shapes.AddLine(x - 3 * cell, y - cell / 2, x - 3 * cell, y + cell / 2);
                    connector.Name = ShapePrefix + "C" + i;
                    Marshal.ReleaseComObject(connector);

                    labels[i, 0] = "Node " + i;
                }

                var shapeTime = watch.ElapsedMilliseconds;
                sheet.Range["AH1"].Resize[nodes, 1].Value2 = labels; // one bulk write for all labels
                Log.Info("Rendered " + nodes + " nodes (" + nodes * 3 + " shapes) in " + shapeTime + " ms; labels +"
                         + (watch.ElapsedMilliseconds - shapeTime) + " ms");
            }
            finally
            {
                app.Calculation = calculation;
            }
        });

        public static void ClearRendered() => ExcelGateway.Post("Clear rendered", app =>
        {
            if (!(app.ActiveSheet is Worksheet sheet)) return;
            var watch = Stopwatch.StartNew();
            var count = DeleteOwnedShapes(sheet);
            Log.Info("Deleted " + count + " owned shapes in " + watch.ElapsedMilliseconds + " ms (other shapes kept)");
        });

        // ---- Check 5: Ctrl+Z through Application.OnUndo ----

        private static string? _undoSheet;
        private static string? _undoAddress;
        private static object? _undoFormula;

        public static void OnUndoTest() => ExcelGateway.Post("OnUndo test", app =>
        {
            if (!(app.ActiveCell is Range cell))
            {
                Log.Warn("Select a cell first.");
                return;
            }

            _undoSheet = cell.Worksheet.Name;
            _undoAddress = cell.Address;
            _undoFormula = cell.Formula;
            cell.Value2 = "DTS spike " + DateTime.Now.ToString("HH:mm:ss");

            // Must be the last change in the command, otherwise Excel drops it.
            app.OnUndo("Undo DTS spike edit", "DTS_SpikeUndo");
            Log.Info("Wrote " + _undoSheet + "!" + _undoAddress + ". Press Ctrl+Z (or Undo) now; it should restore the old value.");
        }, fastMode: false);

        [ExcelCommand(Name = "DTS_SpikeUndo")]
        public static void SpikeUndo()
        {
            ExcelGateway.RunNow("Spike undo", app =>
            {
                if (_undoSheet == null || _undoAddress == null) return;
                var sheet = (Worksheet)app.ActiveWorkbook.Worksheets[_undoSheet];
                sheet.Range[_undoAddress].Formula = _undoFormula;
                Log.Info("Ctrl+Z ran DTS_SpikeUndo: restored " + _undoSheet + "!" + _undoAddress);
                _undoSheet = null;
            });
        }

        // ---- Check 7: a command clicked while a cell is being edited is deferred ----

        public static void DeferredCommand()
        {
            Log.Info("Deferred command clicked at " + DateTime.Now.ToString("HH:mm:ss.fff")
                     + ". If you are editing a cell it runs only after you press Enter/Esc.");
            ExcelGateway.Post("Deferred command", app =>
            {
                Log.Info("Deferred command ran at " + DateTime.Now.ToString("HH:mm:ss.fff"));
            });
        }

        // ---- Macro entry points so tools can drive the checks with Application.Run ----

        [ExcelCommand(Name = "DTS_SpikeWriteXml")] public static void CmdWriteXml() => WriteXmlPart();
        [ExcelCommand(Name = "DTS_SpikeReadXml")] public static void CmdReadXml() => ReadXmlPart();
        [ExcelCommand(Name = "DTS_SpikeRender500")] public static void CmdRender500() => Render(500);
        [ExcelCommand(Name = "DTS_SpikeRender1000")] public static void CmdRender1000() => Render(1000);
        [ExcelCommand(Name = "DTS_SpikeClear")] public static void CmdClear() => ClearRendered();

        // ---- helpers ----

        private static Workbook? RequireWorkbook(Application app)
        {
            var wb = app.ActiveWorkbook;
            if (wb == null) Log.Warn("Open a workbook first.");
            return wb;
        }

        private static Worksheet FindOrAddSheet(Workbook wb)
        {
            foreach (Worksheet ws in wb.Worksheets)
                if (ws.Name == SheetName) return ws;

            var added = (Worksheet)wb.Worksheets.Add(After: wb.Worksheets[wb.Worksheets.Count]);
            added.Name = SheetName;
            return added;
        }

        /// <summary>Delete only shapes this add-in created; everything else on the sheet survives.</summary>
        private static int DeleteOwnedShapes(Worksheet sheet)
        {
            var shapes = sheet.Shapes;
            var deleted = 0;
            for (var i = shapes.Count; i >= 1; i--)
            {
                var shape = shapes.Item(i);
                if (shape.Name.StartsWith(ShapePrefix, StringComparison.Ordinal))
                {
                    shape.Delete();
                    deleted++;
                }
                Marshal.ReleaseComObject(shape);
            }
            return deleted;
        }
    }
}
