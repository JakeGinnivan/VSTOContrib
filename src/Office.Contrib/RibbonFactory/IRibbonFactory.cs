using System;
using System.Reflection;
using Microsoft.Office.Core;
using Microsoft.Office.Tools;

namespace Office.Contrib.RibbonFactory
{
    /// <summary>
    /// Instance of a Ribbon Factory
    /// </summary>
    public interface IRibbonFactory : IRibbonExtensibility
    {
        /// <summary>
        /// Initialises and builds up the ribbon factory
        /// </summary>
        /// <param name="ribbonFactory">The ribbon factory.</param>
        /// <param name="customTaskPaneCollection">The custom task pane collection.</param>
        /// <param name="assemblies">The assemblies to scan for view models.</param>
        /// <returns>
        /// Disposible object to call on outlook shutdown
        /// </returns>
        /// <exception cref="ViewNotFoundException">If the view cannot be located for a view model</exception>
        IDisposable InitialiseFactory(
            Func<Type, IRibbonViewModel> ribbonFactory, 
            CustomTaskPaneCollection customTaskPaneCollection, 
            params Assembly[] assemblies);

        /// <summary>
        /// Gets or sets the locate view strategy.
        /// </summary>
        /// <value>The locate view strategy.</value>
        ViewLocationStrategyBase LocateViewStrategy { get; set; }
    }
}