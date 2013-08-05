namespace VSTOContrib.Core.RibbonFactory.Interfaces.Internal
{
    internal interface ICustomTaskPaneRegister
    {
        void RegisterCustomTaskPanes(IRibbonViewModel ribbonViewModel, object view, object viewContext);
        void Cleanup(object view);
        void ChangeVisibilityForContext(object context, bool visible);
    }
}