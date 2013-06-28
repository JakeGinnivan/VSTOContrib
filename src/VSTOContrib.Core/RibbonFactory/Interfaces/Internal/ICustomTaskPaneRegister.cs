using Microsoft.Office.Tools;

namespace VSTOContrib.Core.RibbonFactory.Interfaces.Internal
{
    internal interface ICustomTaskPaneRegister
    {
        void RegisterCustomTaskPanes(IRibbonViewModel ribbonViewModel, object view);
        void Cleanup(object view);
    }
}