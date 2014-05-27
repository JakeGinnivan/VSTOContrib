using System;

namespace VSTOContrib.Core.RibbonFactory.Interfaces
{
    public interface IViewProvider : IDisposable
    {
        void Initialise();
        event EventHandler<NewViewEventArgs> NewView;
        event EventHandler<ViewClosedEventArgs> ViewClosed;
        void CleanupReferencesTo(object view, object context);
    }
}