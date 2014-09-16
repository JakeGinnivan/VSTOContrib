using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using CustomTaskPane = Microsoft.Office.Tools.CustomTaskPane;

namespace VSTOContrib.Core.RibbonFactory.Internal
{
    class OneToManyCustomTaskPaneAdapter : ICustomTaskPaneWrapper
    {
        readonly Dictionary<OfficeWin32Window, CustomTaskPane> customTaskPanes;
        readonly string title;
        bool disposed;
        bool hasBeenHidden;
        bool visible;
        int width;
        int height;
        MsoCTPDockPosition dockPosition;
        MsoCTPDockPositionRestrict dockPositionRestrict;

        public OneToManyCustomTaskPaneAdapter(string title)
        {
            this.title = title;
            customTaskPanes = new Dictionary<OfficeWin32Window, CustomTaskPane>();
        }

        public void Add(OfficeWin32Window window, CustomTaskPane customTaskPane)
        {
            if (disposed) return;
            if (customTaskPanes.Count == 0)
            {
                visible = customTaskPane.Visible;
                dockPosition = customTaskPane.DockPosition;
                width = customTaskPane.Width;
                height = customTaskPane.Height;
                dockPositionRestrict = customTaskPane.DockPositionRestrict;
            }
            else
            {
                //Sync new task pane's properties up
                customTaskPane.Visible = visible;
                customTaskPane.DockPosition = dockPosition;

                if (dockPosition != MsoCTPDockPosition.msoCTPDockPositionTop &&
                    dockPosition != MsoCTPDockPosition.msoCTPDockPositionBottom)
                {
                    customTaskPane.Width = width;
                }
                if (dockPosition != MsoCTPDockPosition.msoCTPDockPositionLeft &&
                    dockPosition != MsoCTPDockPosition.msoCTPDockPositionRight)
                {
                    customTaskPane.Height = height;
                }
            }

            customTaskPanes.Add(window, customTaskPane);
            customTaskPane.DockPositionChanged += CustomTaskPaneDockPositionChanged;
            customTaskPane.VisibleChanged += CustomTaskPaneVisibleChanged;
        }

        void CustomTaskPaneVisibleChanged(object sender, EventArgs e)
        {
            if (disposed) return;
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
            if (disposed) return;
            var customTaskPane = (CustomTaskPane)sender;
            Do(c => c.DockPositionChanged -= CustomTaskPaneDockPositionChanged);

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
            if (disposed) return;
            foreach (var customTaskPane in customTaskPanes.ToArray())
            {
                action(customTaskPane.Value);
            }
        }

        public string Title
        {
            get { return title; }
        }

        public MsoCTPDockPosition DockPosition
        {
            get { return dockPosition; }
            set
            {
                dockPosition = value;
                Do(c => c.DockPosition = dockPosition);
            }
        }

        public MsoCTPDockPositionRestrict DockPositionRestrict
        {
            get { return dockPositionRestrict; }
            set
            {
                dockPositionRestrict = value;
                Do(c => c.DockPositionRestrict = value);
            }
        }

        public bool Visible
        {
            get { return visible; }
            set
            {
                visible = value;
                Do(c => c.Visible = visible);
            }
        }

        public event EventHandler VisibleChanged;
        public event EventHandler DockPositionChanged;

        public int Height
        {
            get { return height; }
            set
            {
                height = value;
                Do(c => c.Height = height);
            }
        }

        public int Width
        {
            get { return width; }
            set
            {
                width = value;
                Do(c => c.Width = width);
            }
        }

        public void Dispose()
        {
            if (disposed) return;
            Do(DisposeTaskPane);
            disposed = true;
        }

        public void CleanupView(OfficeWin32Window view)
        {
            if (disposed) return;
            var toRemove = customTaskPanes[view];
            DisposeTaskPane(toRemove);
            customTaskPanes.Remove(view);
        }

        void DisposeTaskPane(CustomTaskPane c)
        {
            try
            {
                c.VisibleChanged -= CustomTaskPaneVisibleChanged;
                c.DockPositionChanged -= CustomTaskPaneDockPositionChanged;
                c.Dispose();
            }
            catch (COMException)
            {
            }
            catch (ObjectDisposedException)
            {
            }
        }

        public void HideIfVisible()
        {
            if (Visible)
            {
                Visible = false;
                hasBeenHidden = true;
            }
        }

        public void RestoreIfNeeded()
        {
            if (hasBeenHidden)
            {
                Visible = true;
                hasBeenHidden = false;
            }
        }
    }
}
