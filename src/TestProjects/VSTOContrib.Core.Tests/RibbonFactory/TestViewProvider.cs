using System;
using VSTOContrib.Core.RibbonFactory;
using VSTOContrib.Core.RibbonFactory.Interfaces;

namespace VSTOContrib.Core.Tests.RibbonFactory
{
    public class TestViewProvider : IViewProvider
    {
        public void Dispose()
        {
            throw new NotImplementedException();
        }

        public void Initialise(object application)
        {
            throw new NotImplementedException();
        }

        public event EventHandler<NewViewEventArgs> NewView;
        public event EventHandler<ViewClosedEventArgs> ViewClosed;
        public void CleanupReferencesTo(OfficeWin32Window view, object context)
        {
            throw new NotImplementedException();
        }

        public OfficeWin32Window ToOfficeWindow(object view)
        {
            throw new NotImplementedException();
        }
    }
}