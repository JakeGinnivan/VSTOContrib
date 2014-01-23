using Microsoft.Office.Core;

namespace VSTOContrib.Core.RibbonFactory.Interfaces
{
    internal interface IViewModelResolver
    {
        IRibbonViewModel ResolveInstanceFor(object context);
        void RibbonLoaded(IRibbonUI ribbonUi);
        void RegisterCallbackControl(string ribbonType, string controlCallback, string ribbonControl);
    }
}