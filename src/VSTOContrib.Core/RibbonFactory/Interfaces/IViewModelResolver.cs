using Microsoft.Office.Core;

namespace VSTOContrib.Core.RibbonFactory.Interfaces
{
    internal interface IViewModelResolver<in TRibbonTypes> where TRibbonTypes : struct
    {
        IRibbonViewModel ResolveInstanceFor(object context);
        void RibbonLoaded(IRibbonUI ribbonUi);
        void RegisterCallbackControl(TRibbonTypes ribbonType, string controlCallback, string ribbonControl);
    }
}