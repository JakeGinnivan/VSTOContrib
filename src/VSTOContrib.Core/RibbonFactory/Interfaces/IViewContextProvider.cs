namespace VSTOContrib.Core.RibbonFactory.Interfaces
{
    public interface IViewContextProvider
    {
        object GetContextForView(OfficeWin32Window view);
        string GetRibbonTypeForView(OfficeWin32Window view);
    }
}
