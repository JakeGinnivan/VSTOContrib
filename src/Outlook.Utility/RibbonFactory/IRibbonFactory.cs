using System;
using System.Reflection;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools;
using Office.Utility;

namespace Outlook.Utility.RibbonFactory
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
        /// <param name="outlookApplication">The outlook application.</param>
        /// <param name="customTaskPaneCollection">The custom task pane collection.</param>
        /// <param name="assemblies">The assemblies to scan for view models.</param>
        /// <returns>
        /// Disposible object to call on outlook shutdown
        /// </returns>
        /// <exception cref="ViewNotFoundException">If the view cannot be located for a view model</exception>
        IDisposable InitialiseFactory(
            Func<Type, IRibbonViewModel> ribbonFactory, 
            Application outlookApplication,
            CustomTaskPaneCollection customTaskPaneCollection, 
            params Assembly[] assemblies);

        /// <summary>
        /// Gets or sets the locate view strategy.
        /// </summary>
        /// <value>The locate view strategy.</value>
        ViewLocationStrategyBase LocateViewStrategy { get; set; }
    }
}