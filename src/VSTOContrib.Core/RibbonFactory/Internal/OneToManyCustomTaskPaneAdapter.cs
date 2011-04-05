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

        public OneToManyCustomTaskPaneAdapter(CustomTaskPane original)
        {
            _original = original;
            _customTaskPanes = new List<CustomTaskPane>();
            Add(original);
        }

        public bool ViewRegistered(object view)
        {
            return _customTaskPanes.Any(c => c.Window == view);
        }

        public void Add(CustomTaskPane customTaskPane)
        {
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
            var customTaskPane = (CustomTaskPane) sender;
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
            Do(c => c.VisibleChanged -= CustomTaskPaneVisibleChanged);
            Do(c => c.DockPositionChanged -= CustomTaskPaneDockPositionChanged);
            Do(c => c.Dispose());
        }

        public void CleanupView(object view)
        {
            foreach (var customTaskPane in _customTaskPanes.Where(customTaskPane => customTaskPane.Window == view))
            {
                _customTaskPanes.Remove(customTaskPane);
                customTaskPane.Dispose();
                break;
            }
        }
    }

    /// <summary>
    /// Office 2007 CustomTaskPane is a sealed class, Office2010 is a interface.
    /// This wrapper interface is so I can maintain compatibility with both
    /// </summary>
    public interface ICustomTaskPaneWrapper : IDisposable
    {
        /// <summary>
        /// Gets the control.
        /// </summary>
        /// <value>The control.</value>
        UserControl Control { get; }
        /// <summary>
        /// Gets the title.
        /// </summary>
        /// <value>The title.</value>
        string Title { get; }
        /// <summary>
        /// Gets the window.
        /// </summary>
        /// <value>The window.</value>
        object Window { get; }
        /// <summary>
        /// Gets or sets the dock position.
        /// </summary>
        /// <value>The dock position.</value>
        Microsoft.Office.Core.MsoCTPDockPosition DockPosition { get; set; }
        /// <summary>
        /// Gets or sets the dock position restrict.
        /// </summary>
        /// <value>The dock position restrict.</value>
        Microsoft.Office.Core.MsoCTPDockPositionRestrict DockPositionRestrict { get; set; }
        /// <summary>
        /// Gets or sets the width.
        /// </summary>
        /// <value>The width.</value>
        int Width { get; set; }
        /// <summary>
        /// Gets or sets the height.
        /// </summary>
        /// <value>The height.</value>
        int Height { get; set; }
        /// <summary>
        /// Gets or sets a value indicating whether this <see cref="ICustomTaskPaneWrapper"/> is visible.
        /// </summary>
        /// <value><c>true</c> if visible; otherwise, <c>false</c>.</value>
        bool Visible { get; set; }
        /// <summary>
        /// Occurs when [visible changed].
        /// </summary>
        event EventHandler VisibleChanged;
        /// <summary>
        /// Occurs when [dock position changed].
        /// </summary>
        event EventHandler DockPositionChanged;
    }
}
