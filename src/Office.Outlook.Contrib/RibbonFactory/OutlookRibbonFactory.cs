using System;
using System.Reflection;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools;
using Office.Contrib.RibbonFactory;

namespace Office.Outlook.Contrib.RibbonFactory
{
    /// <summary>
    /// Simplifies adding custom Ribbon's to Office. 
    /// Allows the custom Ribbon xml to be wired up to IRibbonViewModel's
    /// by convention. Simply name the Ribbon.xml the same as the ribbon view model class
    /// in the same assembly
    /// </summary>
    public class OutlookRibbonFactory : RibbonFactory<OutlookRibbonType>
    {
        private static _Application _outlookApplication;

        /// <summary>
        /// Initializes a new instance of the <see cref="OutlookRibbonFactory"/> class.
        /// </summary>
        /// <param name="viewLocationStrategy">The view location strategy, null for default.</param>
        public OutlookRibbonFactory(IViewLocationStrategy viewLocationStrategy = null)
            : base(viewLocationStrategy)
        {
            
        }

        /// <summary>
        /// Initialises the factory.
        /// </summary>
        /// <param name="ribbonFactory">The ribbon factory.</param>
        /// <param name="customTaskPaneCollection">The custom task pane collection.</param>
        /// <param name="assemblies">The assemblies.</param>
        /// <returns></returns>
        public override IDisposable InitialiseFactory(
            Func<Type, IRibbonViewModel> ribbonFactory, 
            CustomTaskPaneCollection customTaskPaneCollection, 
            params Assembly[] assemblies)
        {
            if (_outlookApplication == null)
                throw new InvalidOperationException("Set Outlook application instance first trough SetApplication()");

            return base.InitialiseFactory(ribbonFactory, customTaskPaneCollection, assemblies);
        }

        /// <summary>
        /// The Outlook View Provider, which knows about both Explorers and Inspectors
        /// </summary>
        /// <returns></returns>
        protected override IViewProvider<OutlookRibbonType> ViewProvider()
        {
            return new OutlookViewProvider(_outlookApplication);
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