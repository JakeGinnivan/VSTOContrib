using System;

namespace VSTOContrib.Core.RibbonFactory.Interfaces
{
    internal interface ICustomTaskPaneRegister : IDisposable
    {
        void RegisterCustomTaskPanes(IRibbonViewModel ribbonViewModel, object view, object viewContext);
        void Cleanup(object view);
        void ChangeVisibilityForContext(object context, bool visible);
        void CleanupViewModel(IRibbonViewModel viewModelInstance);
    }
}