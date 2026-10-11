using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using DecisionTree.Excel.Diagnostics;
using DecisionTree.Excel.Gateway;
using ExcelDna.Integration.CustomUI;
using Microsoft.Office.Interop.Excel;

namespace DecisionTree.Excel.UI
{
    /// <summary>
    /// Excel shows one window per workbook, and a custom task pane belongs to one window.
    /// Panes are kept per window handle and removed when their workbook closes.
    /// While any pane is visible, the selection is polled so clicks on shapes (which raise
    /// no Excel event) are noticed too.
    /// </summary>
    internal static class PaneManager
    {
        private const string Title = "DecisionTree Studio";

        private static readonly Dictionary<int, CustomTaskPane> Panes = new Dictionary<int, CustomTaskPane>();
        private static System.Windows.Forms.Timer? _poll;
        private static string _lastSelection = "";

        public static void Initialize()
        {
            _poll = new System.Windows.Forms.Timer { Interval = 300 };
            _poll.Tick += (s, e) => PollSelection();
            _poll.Start();
        }

        public static void Shutdown()
        {
            _poll?.Stop();
            _poll?.Dispose();
            _poll = null;
            foreach (var pane in Panes.Values)
            {
                try { pane.Delete(); } catch { }
            }
            Panes.Clear();
        }

        /// <summary>Show or hide the pane for the active window. Call from a macro context.</summary>
        public static void ToggleForActiveWindow()
        {
            var window = ExcelGateway.App.ActiveWindow;
            if (window == null)
            {
                Log.Warn("Open a workbook before showing the panel.");
                return;
            }

            var hwnd = window.Hwnd;
            if (!Panes.TryGetValue(hwnd, out var pane))
            {
                pane = CustomTaskPaneFactory.CreateCustomTaskPane(typeof(PaneHost), Title, window);
                pane.Width = 360;
                Panes[hwnd] = pane;
                Log.Info("Created pane for window " + hwnd + " (" + Panes.Count + " pane(s) open)");
            }

            pane.Visible = !pane.Visible;
        }

        /// <summary>Remove panes whose windows belong to a closing workbook.</summary>
        public static void OnWorkbookBeforeClose(Workbook workbook)
        {
            foreach (Window window in workbook.Windows)
            {
                var hwnd = window.Hwnd;
                if (Panes.TryGetValue(hwnd, out var pane))
                {
                    try { pane.Delete(); } catch { }
                    Panes.Remove(hwnd);
                    Log.Info("Removed pane for window " + hwnd);
                }
            }
        }

        private static void PollSelection()
        {
            if (!Panes.Values.Any(p => SafeVisible(p))) return;

            string text;
            try
            {
                text = DescribeSelection(ExcelGateway.App.Selection);
            }
            catch (COMException)
            {
                // Excel is busy (e.g. cell-edit mode); try again on the next tick.
                return;
            }
            catch (Exception ex)
            {
                text = "(error: " + ex.Message + ")";
            }

            if (text == _lastSelection) return;
            _lastSelection = text;

            foreach (var pane in Panes.Values)
            {
                if (pane.ContentControl is PaneHost host) host.Panel.SetSelection(text);
            }
        }

        private static string DescribeSelection(object? selection)
        {
            switch (selection)
            {
                case null:
                    return "(none)";
                case Range range:
                    return "cells " + range.Worksheet.Name + "!" + range.Address[false, false];
                default:
                    // Shapes come through as Oval, Rectangle, Line, DrawingObjects... all expose ShapeRange.
                    try
                    {
                        dynamic shapes = ((dynamic)selection).ShapeRange;
                        return "shape " + (string)shapes.Item(1).Name + (shapes.Count > 1 ? " (+" + (shapes.Count - 1) + ")" : "");
                    }
                    catch
                    {
                        return Microsoft.VisualBasic.Information.TypeName(selection);
                    }
            }
        }

        private static bool SafeVisible(CustomTaskPane pane)
        {
            try { return pane.Visible; } catch { return false; }
        }
    }
}
