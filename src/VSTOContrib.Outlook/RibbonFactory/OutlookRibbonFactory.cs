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
        OutlookViewProvider viewProvider;

        /// <summary>
        /// Initializes a new instance of the <see cref="OutlookRibbonFactory"/> class.
        /// </summary>
        /// <param name="ribbonFactory">A delegate taking a type and returning an instance of the requested type</param>
        /// <param name="customTaskPaneCollection">A delayed resolution instance of the custom task pane collection of your addin 'new Lazy(()=>CustomTaskPaneCollection)'</param>
        /// <param name="vstoFactory">The VSTO factory (Globals.Factory)</param>
        /// <param name="assemblies">Assemblies to scan for view models</param>
        public OutlookRibbonFactory(
            Func<Type, IRibbonViewModel> ribbonFactory,
            Lazy<CustomTaskPaneCollection> customTaskPaneCollection,
            Factory vstoFactory, 
            params Assembly[] assemblies)
            : base(new RibbonFactoryController<OutlookRibbonType>(assemblies, new OutlookViewContextProvider(), ribbonFactory, customTaskPaneCollection, vstoFactory))
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="OutlookRibbonFactory"/> class.
        /// </summary>
        /// <param name="ribbonFactory">A delegate taking a type and returning an instance of the requested type</param>
        /// <param name="customTaskPaneCollection">A delayed resolution instance of the custom task pane collection of your addin 'new Lazy(()=>CustomTaskPaneCollection)'</param>
        /// <param name="assemblies">Assemblies to scan for view models</param>
        /// <param name="vstoFactory">The VSTO factory (Globals.Factory)</param>
        /// <param name="viewLocationStrategy">The view location strategy, null for default strategy.</param>
        public OutlookRibbonFactory(
            Func<Type, IRibbonViewModel> ribbonFactory,
            Lazy<CustomTaskPaneCollection> customTaskPaneCollection,
            IViewLocationStrategy viewLocationStrategy,
            Factory vstoFactory, 
            params Assembly[] assemblies)
            : base(new RibbonFactoryController<OutlookRibbonType>(assemblies, new OutlookViewContextProvider(), ribbonFactory, customTaskPaneCollection, vstoFactory, viewLocationStrategy))
        {
        }

        protected override void InitialiseRibbonFactoryController(IRibbonFactoryController controller, object application)
        {
            viewProvider = new OutlookViewProvider((_Application) application);

            controller.Initialise(viewProvider);
        }
        protected override void ShuttingDown()
        {
            viewProvider.Dispose();
        }
    }
}