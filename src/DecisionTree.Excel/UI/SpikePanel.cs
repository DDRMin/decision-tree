using System;
using System.Windows;
using System.Windows.Controls;
using DecisionTree.Excel.Diagnostics;
using DecisionTree.Excel.Spike;

namespace DecisionTree.Excel.UI
{
    /// <summary>
    /// Phase 0 WPF panel: shows the current selection and a log, and runs the spike checks.
    /// Built in code (no XAML) to keep the spike small. Replaced by the real editor in Phase 2.
    /// </summary>
    internal sealed class SpikePanel : UserControl
    {
        private const int MaxLogLines = 300;

        private readonly TextBlock _selection = new TextBlock { TextWrapping = TextWrapping.Wrap, Margin = new Thickness(0, 0, 0, 8) };
        private readonly ListBox _log = new ListBox { FontFamily = new System.Windows.Media.FontFamily("Consolas"), FontSize = 11 };

        public SpikePanel()
        {
            var buttons = new WrapPanel { Margin = new Thickness(0, 0, 0, 8) };
            AddButton(buttons, "Write XML part", SpikeChecks.WriteXmlPart);
            AddButton(buttons, "Read XML part", SpikeChecks.ReadXmlPart);
            AddButton(buttons, "Render 1000 nodes", () => SpikeChecks.Render(1000));
            AddButton(buttons, "Render 500 nodes", () => SpikeChecks.Render(500));
            AddButton(buttons, "Clear rendered", SpikeChecks.ClearRendered);
            AddButton(buttons, "OnUndo test", SpikeChecks.OnUndoTest);
            AddButton(buttons, "Deferred command", SpikeChecks.DeferredCommand);

            var root = new DockPanel { Margin = new Thickness(8) };
            var header = new StackPanel();
            header.Children.Add(new TextBlock { Text = "Phase 0 spike", FontWeight = FontWeights.SemiBold, FontSize = 14, Margin = new Thickness(0, 0, 0, 6) });
            header.Children.Add(_selection);
            header.Children.Add(buttons);
            header.Children.Add(new TextBlock { Text = "Log (" + Log.Directory + ")", TextWrapping = TextWrapping.Wrap, Foreground = System.Windows.Media.Brushes.Gray });
            DockPanel.SetDock(header, Dock.Top);
            root.Children.Add(header);
            root.Children.Add(_log);
            Content = root;

            SetSelection("(nothing yet)");
            Log.Written += OnLogWritten;
            Unloaded += (s, e) => Log.Written -= OnLogWritten;
        }

        public void SetSelection(string text) => _selection.Text = "Selection: " + text;

        private static void AddButton(WrapPanel panel, string text, Action action)
        {
            var button = new Button { Content = text, Margin = new Thickness(0, 0, 6, 6), Padding = new Thickness(8, 3, 8, 3) };
            button.Click += (s, e) => action();
            panel.Children.Add(button);
        }

        private void OnLogWritten(string line)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                _log.Items.Add(line);
                while (_log.Items.Count > MaxLogLines) _log.Items.RemoveAt(0);
                _log.ScrollIntoView(line);
            }));
        }
    }
}
