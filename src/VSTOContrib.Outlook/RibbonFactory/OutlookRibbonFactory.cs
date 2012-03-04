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
        private static _Application _outlookApplication;

        /// <summary>
        /// Initializes a new instance of the <see cref="OutlookRibbonFactory"/> class.
        /// </summary>
        public OutlookRibbonFactory(
            params Assembly[] assemblies)
            : base(new RibbonFactoryController<OutlookRibbonType>(assemblies))
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="OutlookRibbonFactory"/> class.
        /// </summary>
        public OutlookRibbonFactory(
            IViewLocationStrategy viewLocationStrategy,
            params Assembly[] assemblies)
            : base(new RibbonFactoryController<OutlookRibbonType>(assemblies, viewLocationStrategy))
        {
        }

        /// <summary>
        /// Initialises the factory.
        /// </summary>
        /// <param name="ribbonFactory">The ribbon factory.</param>
        /// <param name="customTaskPaneCollection">The custom task pane collection.</param>
        /// <returns>
        /// Disposible object to call on outlook shutdown
        /// </returns>
        /// <exception cref="ViewNotFoundException">If the view cannot be located for a view model</exception>
        public override IDisposable InitialiseFactory(
            Func<Type, IRibbonViewModel> ribbonFactory,
            CustomTaskPaneCollection customTaskPaneCollection)
        {
            if (_outlookApplication == null)
                throw new InvalidOperationException("Set Outlook application instance first trough SetApplication()");

            return InitialiseFactoryInternal(
                new OutlookViewProvider(_outlookApplication), 
                ribbonFactory,
                new OutlookViewContextProvider(),
                customTaskPaneCollection);
        }

        /// <summary>
        /// Sets the Outlook application Instance
        /// </summary>
        /// <param name="application"></param>
        public static void SetApplication(_Application application)
        {
            _outlookApplication = application;
        }
    }
}