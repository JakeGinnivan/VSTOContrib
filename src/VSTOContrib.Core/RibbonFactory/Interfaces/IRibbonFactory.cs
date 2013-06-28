using System;
using Microsoft.Office.Core;
using Microsoft.Office.Tools;

namespace VSTOContrib.Core.RibbonFactory.Interfaces
{
    /// <summary>
    /// Instance of a Ribbon Factory
    /// </summary>
    public interface IRibbonFactory : IRibbonExtensibility
    {
        /// <summary>
        /// Initialises and builds up the ribbon factory
        /// </summary>
        /// <param name="customTaskPaneCollection">The custom task pane collection.</param>
        /// <returns>
        /// Disposible object to call on outlook shutdown
        /// </returns>
        /// <exception cref="ViewNotFoundException">If the view cannot be located for a view model</exception>
        IDisposable InitialiseFactory(
            CustomTaskPaneCollection customTaskPaneCollection);

        /// <summary>
        /// Gets or sets the locate view strategy.
        /// </summary>
        /// <value>The locate view strategy.</value>
        IViewLocationStrategy LocateViewStrategy { get; set; }
    }
}