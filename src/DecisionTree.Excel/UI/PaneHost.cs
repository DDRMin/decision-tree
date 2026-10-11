using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Windows.Forms.Integration;

namespace DecisionTree.Excel.UI
{
    /// <summary>WinForms control that Excel hosts as a custom task pane; it wraps the WPF panel.</summary>
    [ComVisible(true)]
    public class PaneHost : UserControl
    {
        internal SpikePanel Panel { get; }

        public PaneHost()
        {
            Panel = new SpikePanel();
            Controls.Add(new ElementHost { Dock = DockStyle.Fill, Child = Panel });
        }
    }
}
