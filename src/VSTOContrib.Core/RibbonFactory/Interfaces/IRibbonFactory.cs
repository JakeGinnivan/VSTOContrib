using Microsoft.Office.Core;

namespace VSTOContrib.Core.RibbonFactory.Interfaces
{
    /// <summary>
    /// Instance of a Ribbon Factory
    /// </summary>
    public interface IRibbonFactory : IRibbonExtensibility
    {
        /// <summary>
        /// Gets the locate view strategy.
        /// </summary>
        /// <value>The locate view strategy.</value>
        IViewLocationStrategy LocateViewStrategy { get; }

        /// <summary>
        /// Gets or sets the view model factory, default uses Activator.CreateInstance
        /// </summary>
        IViewModelFactory ViewModelFactory { get; set; }
    }
}