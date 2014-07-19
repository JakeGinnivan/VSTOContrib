using System;

namespace VSTOContrib.Core.RibbonFactory.Interfaces
{
    public interface IViewProvider : IDisposable
    {
        void Initialise(object application);
        event EventHandler<NewViewEventArgs> NewView;
        event EventHandler<ViewClosedEventArgs> ViewClosed;
        void CleanupReferencesTo(OfficeWin32Window view, object context);
        OfficeWin32Window ToOfficeWindow(object view);
    }
}