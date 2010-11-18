using System;
using System.Reflection;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;

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
        /// <param name="assemblies">The assemblies to scan for view models.</param>
        /// <returns>
        /// Disposible object to call on outlook shutdown
        /// </returns>
        /// <exception cref="ViewNotFoundException">If the view cannot be located for a view model</exception>
        IDisposable InitialiseFactory(Func<Type, IRibbonViewModel> ribbonFactory, Application outlookApplication,
                                      params Assembly[] assemblies);
    }
}