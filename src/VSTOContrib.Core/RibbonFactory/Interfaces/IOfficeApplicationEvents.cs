using System;

namespace VSTOContrib.Core.RibbonFactory.Interfaces
{
    public interface IOfficeApplicationEvents : IDisposable
    {
        void Initialise(object application);
        /// <summary>
        /// This needs a better name, it is not really NewView.
        /// It is more a view has been opened, activated or something and it should be evaluated
        /// </summary>
        event Action<NewViewEventArgs> NewView;
        event Action<OfficeWin32Window> ViewClosed;
        event Action<object> ContextClosed;
        OfficeWin32Window ToOfficeWindow(object view);
        OfficeWin32Window ActiveWindow { get; }
    }
}