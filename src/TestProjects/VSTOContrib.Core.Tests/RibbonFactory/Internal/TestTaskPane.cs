using System;
using System.Windows.Forms;
using Microsoft.Office.Core;
using CustomTaskPane = Microsoft.Office.Tools.CustomTaskPane;

namespace VSTOContrib.Core.Tests.RibbonFactory.Internal
{
    public class TestTaskPane : CustomTaskPane
    {
        public void Dispose()
        {
            DisposedCalled++;
        }

        public int DisposedCalled { get; private set; }

        public UserControl Control { get; private set; }
        public string Title { get; private set; }
        public object Window { get; private set; }
        public MsoCTPDockPosition DockPosition { get; set; }
        public MsoCTPDockPositionRestrict DockPositionRestrict { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
        public bool Visible { get; set; }
        public event EventHandler VisibleChanged;
        public event EventHandler DockPositionChanged;
    }
}