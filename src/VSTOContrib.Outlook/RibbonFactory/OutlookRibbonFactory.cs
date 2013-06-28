using System;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools;
using VSTOContrib.Core.RibbonFactory;
using VSTOContrib.Core.RibbonFactory.Interfaces;

namespace VSTOContrib.Outlook.RibbonFactory
{
    /// <summary>
    /// Simplifies adding custom Ribbon's to Office. 
    /// Allows the custom Ribbon xml to be wired up to IRibbonViewModel's
    /// by convention. Simply name the Ribbon.xml the same as the ribbon view model class
    /// in the same assembly
    /// </summary>
    [ComVisible(true)]
    public class OutlookRibbonFactory : Core.RibbonFactory.RibbonFactory
    {
        private static _Application outlookApplication;

        /// <summary>
        /// Initializes a new instance of the <see cref="OutlookRibbonFactory"/> class.
        /// </summary>
        /// <param name="ribbonFactory">A delegate taking a type and returning an instance of the requested type</param>
        /// <param name="customTaskPaneCollection">A delayed resolution instance of the custom task pane collection of your addin 'new Lazy(()=>CustomTaskPaneCollection)'</param>
        /// <param name="assemblies">Assemblies to scan for view models</param>
        public OutlookRibbonFactory(
            Func<Type, IRibbonViewModel> ribbonFactory,
            Lazy<CustomTaskPaneCollection> customTaskPaneCollection,
            params Assembly[] assemblies)
            : base(new RibbonFactoryController<OutlookRibbonType>(assemblies, new OutlookViewContextProvider(), ribbonFactory, customTaskPaneCollection))
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="OutlookRibbonFactory"/> class.
        /// </summary>
        /// <param name="ribbonFactory">A delegate taking a type and returning an instance of the requested type</param>
        /// <param name="customTaskPaneCollection">A delayed resolution instance of the custom task pane collection of your addin 'new Lazy(()=>CustomTaskPaneCollection)'</param>
        /// <param name="assemblies">Assemblies to scan for view models</param>
        /// <param name="viewLocationStrategy">The view location strategy, null for default strategy.</param>
        public OutlookRibbonFactory(
            Func<Type, IRibbonViewModel> ribbonFactory,
            Lazy<CustomTaskPaneCollection> customTaskPaneCollection,
            IViewLocationStrategy viewLocationStrategy,
            params Assembly[] assemblies)
            : base(new RibbonFactoryController<OutlookRibbonType>(assemblies, new OutlookViewContextProvider(), ribbonFactory, customTaskPaneCollection, viewLocationStrategy))
        {
        }

        /// <summary>
        /// Initialises the factory.
        /// </summary>
        /// <param name="customTaskPaneCollection">The custom task pane collection.</param>
        /// <returns>
        /// Disposible object to call on outlook shutdown
        /// </returns>
        /// <exception cref="ViewNotFoundException">If the view cannot be located for a view model</exception>
        public override IDisposable InitialiseFactory(
            CustomTaskPaneCollection customTaskPaneCollection)
        {
            if (outlookApplication == null)
                throw new InvalidOperationException("Set Outlook application instance first trough SetApplication()");

            return InitialiseFactoryInternal(
                new OutlookViewProvider(outlookApplication));
        }

        /// <summary>
        /// Sets the Outlook application Instance
        /// </summary>
        /// <param name="application"></param>
        public static void SetApplication(_Application application)
        {
            outlookApplication = application;
        }
    }
}