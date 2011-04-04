using Microsoft.Office.Tools;

namespace Office.Contrib.RibbonFactory.Interfaces.Internal
{
    internal interface ICustomTaskPaneRegister
    {
        void Initialise(CustomTaskPaneCollection customTaskPaneCollection);
        void RegisterCustomTaskPanes(IRibbonViewModel ribbonViewModel, object view);
        void Cleanup(object view);
    }
}