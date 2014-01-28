using System;

namespace VSTOContrib.Core.RibbonFactory.Interfaces
{
    public interface IViewProvider : IDisposable
    {
        void Initialise();
        event EventHandler<NewViewEventArgs> NewView;
        event EventHandler<ViewClosedEventArgs> ViewClosed;
        event EventHandler<HideCustomTaskPanesForContextEventArgs> UpdateCustomTaskPanesVisibilityForContext;
        void CleanupReferencesTo(object view, object context);
    }
}