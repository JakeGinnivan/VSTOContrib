using System;
using Microsoft.Office.Interop.Outlook;

namespace VSTOContrib.Outlook
{
    public class ExplorerEventArgs : EventArgs
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ExplorerEventArgs"/> class.
        /// </summary>
        /// <param name="explorer">The explorer.</param>
        public ExplorerEventArgs(Explorer explorer)
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