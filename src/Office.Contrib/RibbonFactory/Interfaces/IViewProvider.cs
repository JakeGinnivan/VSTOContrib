using System;

namespace Office.Contrib.RibbonFactory.Interfaces
{
    /// <summary>
    /// 
    /// </summary>
    /// <typeparam name="TRibbonTypes"></typeparam>
    public interface IViewProvider<TRibbonTypes> : IDisposable
    {
        /// <summary>
        /// 
        /// </summary>
        void Initialise();

        /// <summary>
        /// Raise when a new view is created
        /// </summary>
        event EventHandler<NewViewEventArgs<TRibbonTypes>> NewView;

        /// <summary>
        /// Raise when a view is closed, 
        /// </summary>
        event EventHandler<ViewClosedEventArgs> ViewClosed;

        /// <summary>
        /// Unregister any event handlers, and release any references to a view instance
        /// </summary>
        /// <param name="view"></param>
        /// <param name="context"></param>
        void CleanupReferencesTo(object view, object context);
    }
}