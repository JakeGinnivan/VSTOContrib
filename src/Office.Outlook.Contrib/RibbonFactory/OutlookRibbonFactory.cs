using System;
using System.Reflection;
using System.Runtime.InteropServices;
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
    [ComVisible(true)]
    public class OutlookRibbonFactory : Office.Contrib.RibbonFactory.RibbonFactory
    {
        private static _Application _outlookApplication;

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

            return InitialiseFactoryInternal(
                new OutlookViewProvider(_outlookApplication), ribbonFactory, 
                customTaskPaneCollection, assemblies);
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