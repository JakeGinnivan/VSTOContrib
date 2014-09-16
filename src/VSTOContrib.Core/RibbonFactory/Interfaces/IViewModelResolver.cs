using Microsoft.Office.Core;

namespace VSTOContrib.Core.RibbonFactory.Interfaces
{
    internal interface IViewModelResolver
    {
        IRibbonViewModel BuildViewModel(string ribbonType, IRibbonUI ribbonUi, object viewContext);
    }
}