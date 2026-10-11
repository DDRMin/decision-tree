using System;
using DecisionTree.Excel.Diagnostics;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;

namespace DecisionTree.Excel.Gateway
{
    /// <summary>
    /// The only way commands touch the Excel object model.
    /// - Post() queues work with QueueAsMacro, so it runs on Excel's main thread only once
    ///   Excel is ready (e.g. after the user leaves cell-edit mode).
    /// - A reentrancy guard stops a second command starting while one is running.
    /// - ScreenUpdating / EnableEvents are always restored, even after a failure.
    /// </summary>
    internal static class ExcelGateway
    {
        private static bool _running;

        public static Application App => (Application)ExcelDnaUtil.Application;

        /// <summary>Raised on the main thread when a command fails, for the panel to show.</summary>
        public static event Action<string, Exception>? CommandFailed;

        public static void Post(string name, Action<Application> action, bool fastMode = true)
        {
            Log.Info(name + ": queued");
            ExcelAsyncUtil.QueueAsMacro(() => RunNow(name, action, fastMode));
        }

        /// <summary>Run immediately. Only call from a macro context (ribbon callback, QueueAsMacro, ExcelCommand).</summary>
        public static void RunNow(string name, Action<Application> action, bool fastMode = true)
        {
            if (_running)
            {
                Log.Warn(name + ": ignored, another command is still running");
                return;
            }

            _running = true;
            var app = App;
            var screenUpdating = true;
            var enableEvents = true;
            try
            {
                if (fastMode)
                {
                    screenUpdating = app.ScreenUpdating;
                    enableEvents = app.EnableEvents;
                    app.ScreenUpdating = false;
                    app.EnableEvents = false;
                }

                Log.Info(name + ": running");
                action(app);
            }
            catch (Exception ex)
            {
                Log.Error(name + ": failed", ex);
                try { CommandFailed?.Invoke(name, ex); } catch { }
            }
            finally
            {
                // Only write settings back if we changed them: any extra write after
                // Application.OnUndo could drop the registered undo entry.
                if (fastMode)
                {
                    try
                    {
                        app.ScreenUpdating = screenUpdating;
                        app.EnableEvents = enableEvents;
                    }
                    catch (Exception ex)
                    {
                        Log.Error("Restoring Excel settings failed", ex);
                    }
                }

                _running = false;
            }
        }
    }
}
