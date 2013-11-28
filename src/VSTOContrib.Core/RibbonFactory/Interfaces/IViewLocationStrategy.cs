namespace VSTOContrib.Core.RibbonFactory.Interfaces
{
    /// <summary>
    /// View Location Strategy
    /// </summary>
    public interface IViewLocationStrategy
    {
        ///<summary>
        /// Fetches the Ribbon XML for a given view
        ///</summary>
        ///<typeparam name="T">The View model to fetch the Ribbon XML for</typeparam>
        ///<returns>Ribbon XML</returns>
        string LocateViewForViewModel<T>() where T : IRibbonViewModel;
    }
}