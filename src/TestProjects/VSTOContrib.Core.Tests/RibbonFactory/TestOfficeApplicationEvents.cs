using System;
using VSTOContrib.Core.RibbonFactory;
using VSTOContrib.Core.RibbonFactory.Interfaces;

namespace VSTOContrib.Core.Tests.RibbonFactory
{
    public class TestOfficeApplicationEvents : IOfficeApplicationEvents
    {
        public void Dispose()
        {
            throw new NotImplementedException();
        }

        public void Initialise(object application)
        {
            throw new NotImplementedException();
        }

        public event Action<NewViewEventArgs> NewView;
        public event Action<OfficeWin32Window> ViewClosed;
        public event Action<object> ContextClosed;

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