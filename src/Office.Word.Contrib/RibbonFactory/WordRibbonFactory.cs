using System;
using System.Reflection;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools;
using Office.Contrib.RibbonFactory;

namespace Office.Word.Contrib.RibbonFactory
{
    /// <summary>
    /// 
    /// </summary>
    public class WordRibbonFactory : RibbonFactory<WordRibbonType>
    {
        private static Application _wordApplication;

        public override IDisposable InitialiseFactory(
            Func<Type, IRibbonViewModel> ribbonFactory,
            CustomTaskPaneCollection customTaskPaneCollection,
            params Assembly[] assemblies)
        {
            if (_wordApplication == null)
                throw new InvalidOperationException("Set Word application instance first trough SetApplication()");

            return base.InitialiseFactory(ribbonFactory, customTaskPaneCollection, assemblies);
        }

        protected override IViewProvider<WordRibbonType> ViewProvider()
        {
            return new WordViewProvider(_wordApplication);
        }

        /// <summary>
        /// Sets the Outlook application Instance
        /// </summary>
        /// <param name="application"></param>
        public static void SetApplication(Application application)
        {
            _wordApplication = application;
        }
    }
}