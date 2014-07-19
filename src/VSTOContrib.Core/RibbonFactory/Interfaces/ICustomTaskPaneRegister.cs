using System;

namespace VSTOContrib.Core.RibbonFactory.Interfaces
{
    internal interface ICustomTaskPaneRegister : IDisposable
    {
        void RegisterCustomTaskPanes(IRibbonViewModel ribbonViewModel, OfficeWin32Window view, object viewContext);
        void Cleanup(OfficeWin32Window view);
        void CleanupViewModel(IRibbonViewModel viewModelInstance);
    }
}