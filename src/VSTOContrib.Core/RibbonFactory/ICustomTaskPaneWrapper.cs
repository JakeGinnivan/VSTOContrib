using System;
using System.Windows.Forms;

namespace VSTOContrib.Core.RibbonFactory
{
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