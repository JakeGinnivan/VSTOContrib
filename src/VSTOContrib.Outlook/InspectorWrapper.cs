using System;
using Microsoft.Office.Interop.Outlook;

namespace VSTOContrib.Outlook
{
    /// <summary>
    /// 
    /// </summary>
    public class InspectorWrapper
    {
        private object _currentContext;

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

        public object CurrentContext
        {
            get { return _currentContext; }
            set { _currentContext = value; }
        }

        private void InspectorClose()
        {
            ((InspectorEvents_10_Event)Inspector).Close -= InspectorClose;

            var handler = Closed;
            if (handler != null) 
                Closed(this, new InspectorClosedEventArgs(Inspector));

            Inspector = null;
        }
    }
}
