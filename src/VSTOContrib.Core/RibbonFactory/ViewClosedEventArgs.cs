using System;

namespace VSTOContrib.Core.RibbonFactory
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
        public OfficeWin32Window View { get; private set; }

        /// <summary>
        /// Gets or sets the context.
        /// </summary>
        /// <value>The context.</value>
        public object Context { get; private set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="ViewClosedEventArgs"/> class.
        /// </summary>
        /// <param name="view">The view.</param>
        /// <param name="context">The context.</param>
        public ViewClosedEventArgs(OfficeWin32Window view, object context)
        {
            View = view;
            Context = context;
        }
    }
}