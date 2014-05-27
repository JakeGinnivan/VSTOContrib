using System;

namespace VSTOContrib.Core.RibbonFactory
{
    public class HideCustomTaskPanesForContextEventArgs : EventArgs
    {
        public HideCustomTaskPanesForContextEventArgs(object context, bool visible)
        {
            Context = context;
            Visible = visible;
        }

        /// <summary>
        /// The context which the visibility should be toggled for
        /// </summary>
        public object Context { get; private set; }

        /// <summary>
        /// False to hide any open task pane, true to make visible (or restore existing value)
        /// </summary>
        public bool Visible { get; private set; }
    }
}