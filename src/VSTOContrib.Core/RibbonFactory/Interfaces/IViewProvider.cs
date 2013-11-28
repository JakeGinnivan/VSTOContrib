using System;

namespace VSTOContrib.Core.RibbonFactory.Interfaces
{
    public interface IViewProvider : IDisposable
    {
        /// <summary>
        /// 
        /// </summary>
        void Initialise();

        /// <summary>
        /// Raise when a new view is created
        /// </summary>
        event EventHandler<NewViewEventArgs> NewView;

        /// <summary>
        /// Raise when a view is closed, 
        /// </summary>
        event EventHandler<ViewClosedEventArgs> ViewClosed;

        /// <summary>
        /// Raise when the custom task panes for a context need to change their visibility
        /// </summary>
        event EventHandler<HideCustomTaskPanesForContextEventArgs> UpdateCustomTaskPanesVisibilityForContext;

        /// <summary>
        /// Unregister any event handlers, and release any references to a view instance
        /// </summary>
        /// <param name="view"></param>
        /// <param name="context"></param>
        void CleanupReferencesTo(object view, object context);
    }
}