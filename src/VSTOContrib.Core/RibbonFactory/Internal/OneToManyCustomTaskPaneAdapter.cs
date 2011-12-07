using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Tools;

namespace VSTOContrib.Core.RibbonFactory.Internal
{
    class OneToManyCustomTaskPaneAdapter : ICustomTaskPaneWrapper
    {
        private readonly CustomTaskPane _original;
        private readonly List<CustomTaskPane> _customTaskPanes;
        private bool _disposed;

        public OneToManyCustomTaskPaneAdapter(CustomTaskPane original)
        {
            _original = original;
            _customTaskPanes = new List<CustomTaskPane>();
            Add(original);
        }

        public bool ViewRegistered(object view)
        {
            if (_disposed) return false;
            return _customTaskPanes.Any(c => c.Window == view);
        }

        public void Add(CustomTaskPane customTaskPane)
        {
            if (_disposed) return;
            //Sync new task pane's properties up
            customTaskPane.Visible = _original.Visible;
            customTaskPane.DockPosition = _original.DockPosition;


            if (_original.DockPosition != Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionTop &&
                _original.DockPosition != Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionBottom)
            {
                customTaskPane.Width = _original.Width;
            }
            if (_original.DockPosition != Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionLeft &&
                _original.DockPosition != Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight)
            {
                customTaskPane.Height = _original.Height;
            }
            
            _customTaskPanes.Add(customTaskPane);
            customTaskPane.DockPositionChanged += CustomTaskPaneDockPositionChanged;
            customTaskPane.VisibleChanged += CustomTaskPaneVisibleChanged;
        }

        public void Refresh(object view)
        {

        }

        void CustomTaskPaneVisibleChanged(object sender, EventArgs e)
        {
            if (_disposed) return;
            var customTaskPane = (CustomTaskPane)sender;
            Do(c => c.VisibleChanged -= CustomTaskPaneVisibleChanged);

            //Propagate changes, then raise adapter event
            Do(c =>
                   {
                       if (c != customTaskPane)
                           c.Visible = customTaskPane.Visible;
                   });
            var handler = VisibleChanged;
            if (handler != null)
                handler(this, EventArgs.Empty);

            Do(c => c.VisibleChanged += CustomTaskPaneVisibleChanged);
        }

        void CustomTaskPaneDockPositionChanged(object sender, EventArgs e)
        {
            if (_disposed) return;
            var customTaskPane = (CustomTaskPane)sender;
            Do(c=>c.DockPositionChanged -= CustomTaskPaneDockPositionChanged);

            //Propagate changes, then raise adapter event
            Do(c =>
                   {
                       if (c != customTaskPane)
                           c.DockPosition = customTaskPane.DockPosition;
                   });
            var handler = DockPositionChanged;
            if (handler != null)
                handler(this, EventArgs.Empty);

            Do(c => c.DockPositionChanged += CustomTaskPaneDockPositionChanged);
        }

        private void Do(Action<CustomTaskPane> action)
        {
            if (_disposed) return;
            foreach (var customTaskPane in _customTaskPanes)
            {
                action(customTaskPane);
            }
        }

        public UserControl Control
        {
            get { return _original.Control; }
        }

        public string Title
        {
            get { return _original.Title; }
        }

        public object Window
        {
            get { return _original.Window; }
        }

        public Microsoft.Office.Core.MsoCTPDockPosition DockPosition
        {
            get { return _original.DockPosition; }
            set { Do(c=>c.DockPosition = value); }
        }

        public Microsoft.Office.Core.MsoCTPDockPositionRestrict DockPositionRestrict
        {
            get { return _original.DockPositionRestrict; }
            set { Do(c => c.DockPositionRestrict = value); }
        }

        public bool Visible
        {
            get { return _original.Visible; }
            set { Do(c=>c.Visible = value); }
        }

        public event EventHandler VisibleChanged;
        public event EventHandler DockPositionChanged;

        public int Height
        {
            get { return _original.Height; }
            set { Do(c=>c.Height = value); }
        }

        public int Width
        {
            get { return _original.Width; }
            set { Do(c => c.Width = value); }
        }

        public void Dispose()
        {
            _disposed = true;
            Do(c => c.VisibleChanged -= CustomTaskPaneVisibleChanged);
            Do(c => c.DockPositionChanged -= CustomTaskPaneDockPositionChanged);
            Do(c => c.Dispose());
        }

        public void CleanupView(object view)
        {
            if (_disposed) return;
            foreach (var customTaskPane in _customTaskPanes.Where(customTaskPane => customTaskPane.Window == view))
            {
                _customTaskPanes.Remove(customTaskPane);
                customTaskPane.Dispose();
                break;
            }
        }
    }
}
