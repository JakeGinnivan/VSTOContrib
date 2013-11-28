using System;
using Microsoft.Office.Interop.Outlook;

namespace VSTOContrib.Outlook
{
    /// <summary>
    /// 
    /// </summary>
    public class InspectorClosedEventArgs : EventArgs
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="inspector"></param>
        /// <param name="currentContext">The current context for the inspector</param>
        public InspectorClosedEventArgs(Inspector inspector, object currentContext)
        {
            Inspector = inspector;
            CurrentContext = currentContext;
        }

        /// <summary>
        /// Gets the closed inspector.
        /// </summary>
        /// <value>The inspector.</value>
        public Inspector Inspector { get; private set; }

        /// <summary>
        /// Gets the current context for the inspector
        /// </summary>
        public object CurrentContext { get; set; }
    }
}