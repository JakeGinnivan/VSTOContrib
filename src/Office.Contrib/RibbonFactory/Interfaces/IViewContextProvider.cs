namespace Office.Contrib.RibbonFactory.Interfaces
{
    /// <summary>
    /// 
    /// </summary>
    public interface IViewContextProvider
    {
        /// <summary>
        /// Gets the context for view.
        /// </summary>
        /// <param name="view">The view.</param>
        /// <returns></returns>
        object GetContextForView(object view);
    }
}
