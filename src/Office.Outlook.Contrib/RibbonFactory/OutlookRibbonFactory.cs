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

        public override IDisposable InitialiseFactory(
            Func<Type, IRibbonViewModel> ribbonFactory, 
            CustomTaskPaneCollection customTaskPaneCollection, 
            params Assembly[] assemblies)
        {
            if (_outlookApplication == null)
                throw new InvalidOperationException("Set Outlook application instance first trough SetApplication()");

            return base.InitialiseFactory(ribbonFactory, customTaskPaneCollection, assemblies);
        }

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