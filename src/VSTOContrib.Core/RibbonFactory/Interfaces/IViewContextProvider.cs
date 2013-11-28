namespace VSTOContrib.Core.RibbonFactory.Interfaces
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

        /// <summary>
        /// Gets the ribbon type for view.
        /// </summary>
        /// <typeparam name="TRibbonType">The type of the ribbon type.</typeparam>
        /// <param name="view">The view.</param>
        /// <returns></returns>
        TRibbonType GetRibbonTypeForView<TRibbonType>(object view);
    }
}
