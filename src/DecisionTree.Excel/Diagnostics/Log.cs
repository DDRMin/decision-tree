using System;
using System.IO;

namespace DecisionTree.Excel.Diagnostics
{
    /// <summary>
    /// Local diagnostics log: %LOCALAPPDATA%\DecisionTreeStudio\logs\dts-yyyyMMdd.log.
    /// Never throws; logging must not be able to break an Excel command.
    /// </summary>
    internal static class Log
    {
        private static readonly object Gate = new object();

        public static string Directory { get; } = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "DecisionTreeStudio", "logs");

        /// <summary>Raised for every entry so the panel can show recent messages.</summary>
        public static event Action<string>? Written;

        public static void Info(string message) => Write("INFO", message);

        public static void Warn(string message) => Write("WARN", message);

        public static void Error(string message, Exception? ex = null) =>
            Write("ERROR", ex == null ? message : message + " | " + ex.GetType().Name + ": " + ex.Message);

        private static void Write(string level, string message)
        {
            var line = DateTime.Now.ToString("HH:mm:ss.fff") + " " + level + " " + message;
            try
            {
                lock (Gate)
                {
                    System.IO.Directory.CreateDirectory(Directory);
                    File.AppendAllText(Path.Combine(Directory, "dts-" + DateTime.Now.ToString("yyyyMMdd") + ".log"),
                        line + Environment.NewLine);
                }
            }
            catch
            {
                // Ignore: diagnostics are best effort.
            }

            try { Written?.Invoke(line); } catch { }
        }
    }
}
