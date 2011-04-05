using Microsoft.Office.Tools;

namespace VSTOContrib.Core.RibbonFactory.Interfaces.Internal
{
    internal interface ICustomTaskPaneRegister
    {
        void Initialise(CustomTaskPaneCollection customTaskPaneCollection);
        void RegisterCustomTaskPanes(IRibbonViewModel ribbonViewModel, object view);
        void Cleanup(object view);
    }
}