using System;
using Microsoft.Office.Interop.Outlook;

namespace VSTOContrib.Outlook
{
    /// <summary>
    /// 
    /// </summary>
    public class InspectorWrapper
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="InspectorWrapper"/> class.
        /// </summary>
        /// <param name="inspector">The inspector.</param>
        public InspectorWrapper(Inspector inspector)
        {
            Inspector = inspector;
            ((InspectorEvents_10_Event)Inspector).Close += InspectorClose;
            CurrentContext = Inspector.CurrentItem;
        }

        /// <summary>
        /// Occurs when inspector is closed.
        /// </summary>
        public event EventHandler<InspectorClosedEventArgs> Closed;

        /// <summary>
        /// Gets the inspector.
        /// </summary>
        /// <value>The inspector.</value>
        public Inspector Inspector { get; private set; }

        /// <summary>
        /// Current item for the inspector
        /// </summary>
        public object CurrentContext { get; set; }

        private void InspectorClose()
        {
            ((InspectorEvents_10_Event)Inspector).Close -= InspectorClose;

            var handler = Closed;
            if (handler != null) 
                Closed(this, new InspectorClosedEventArgs(Inspector, CurrentContext));

            Inspector = null;
        }
    }
}
