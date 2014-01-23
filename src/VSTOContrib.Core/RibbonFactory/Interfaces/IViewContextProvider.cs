namespace VSTOContrib.Core.RibbonFactory.Interfaces
{
    public interface IViewContextProvider
    {
        object GetContextForView(object view);
        string GetRibbonTypeForView(object view);
    }
}
