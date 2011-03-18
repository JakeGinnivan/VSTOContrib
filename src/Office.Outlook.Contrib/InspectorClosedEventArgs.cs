using System;
using Microsoft.Office.Interop.Outlook;

namespace Office.Outlook.Contrib
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
        public InspectorClosedEventArgs(Inspector inspector)
        {
            Inspector = inspector;
        }

        /// <summary>
        /// Gets the closed inspector.
        /// </summary>
        /// <value>The inspector.</value>
        public Inspector Inspector { get; private set; }
    }
}