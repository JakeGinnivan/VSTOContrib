using System;
using Microsoft.Office.Interop.Outlook;

namespace Office.Outlook.Contrib
{
    /// <summary>
    /// 
    /// </summary>
    public class ExplorerWrapper
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ExplorerWrapper"/> class.
        /// </summary>
        /// <param name="explorer">The explorer.</param>
        public ExplorerWrapper(Explorer explorer)
        {
            Explorer = explorer;
            ((ExplorerEvents_10_Event)Explorer).Close += ExplorerClose;
        }

        /// <summary>
        /// Occurs when inspector is closed.
        /// </summary>
        public event EventHandler<ExplorerClosedEventArgs> Closed;

        /// <summary>
        /// Gets the inspector.
        /// </summary>
        /// <value>The inspector.</value>
        public Explorer Explorer { get; private set; }

        private void ExplorerClose()
        {
            ((ExplorerEvents_10_Event)Explorer).Close -= ExplorerClose;
            Explorer = null;

            var handler = Closed;
            if (handler != null) 
                Closed(this, new ExplorerClosedEventArgs(Explorer));
        }
    }
}
