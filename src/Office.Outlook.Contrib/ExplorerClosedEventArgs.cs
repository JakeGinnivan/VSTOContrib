using System;
using Microsoft.Office.Interop.Outlook;

namespace Office.Outlook.Contrib
{
    /// <summary>
    /// 
    /// </summary>
    public class ExplorerClosedEventArgs : EventArgs
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ExplorerClosedEventArgs"/> class.
        /// </summary>
        /// <param name="explorer">The explorer.</param>
        public ExplorerClosedEventArgs(Explorer explorer)
        {
            Explorer = explorer;
        }

        /// <summary>
        /// Gets the closed inspector.
        /// </summary>
        /// <value>The inspector.</value>
        public Explorer Explorer { get; private set; }
    }
}