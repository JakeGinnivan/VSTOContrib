using System;

namespace Office.Contrib.RibbonFactory
{
    /// <summary>
    /// 
    /// </summary>
    /// <typeparam name="TRibbonTypes"></typeparam>
    public interface IViewProvider<TRibbonTypes> : IDisposable
    {
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
        void CleanupReferencesTo(object view);
    }
}