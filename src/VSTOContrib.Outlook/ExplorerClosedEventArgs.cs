using System;
using Microsoft.Office.Interop.Outlook;

namespace VSTOContrib.Outlook
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
        /// <param name="currentContext"> </param>
        public ExplorerClosedEventArgs(Explorer explorer, MAPIFolder currentContext)
        {
            Explorer = explorer;
            CurrentContext = currentContext;
        }

        /// <summary>
        /// Gets the closed inspector.
        /// </summary>
        /// <value>The inspector.</value>
        public Explorer Explorer { get; private set; }

        /// <summary>
        /// The currently selected folder on the Explorer window
        /// </summary>
        public MAPIFolder CurrentContext { get; set; }
    }
}