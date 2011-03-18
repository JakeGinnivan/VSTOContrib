using System;

namespace Office.Contrib.RibbonFactory
{
    /// <summary>
    /// 
    /// </summary>
    public class ViewClosedEventArgs : EventArgs
    {
        /// <summary>
        /// Gets the view that was closed.
        /// </summary>
        /// <value>The view.</value>
        public object View { get; private set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="ViewClosedEventArgs"/> class.
        /// </summary>
        /// <param name="view">The view.</param>
        public ViewClosedEventArgs(object view)
        {
            View = view;
        }
    }
}