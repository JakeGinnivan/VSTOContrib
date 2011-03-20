using System;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools;
using Office.Contrib.RibbonFactory;

namespace Office.Word.Contrib.RibbonFactory
{
    /// <summary>
    /// 
    /// </summary>
    [ComVisible(true)]
    public class WordRibbonFactory : Office.Contrib.RibbonFactory.RibbonFactory
    {
        private static Application _wordApplication;

        /// <summary>
        /// Initialises the ribbon factory.
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
            if (_wordApplication == null)
                throw new InvalidOperationException("Set Word application instance first trough SetApplication()");

            return InitialiseFactoryInternal(
                new WordViewProvider(_wordApplication), ribbonFactory, 
                customTaskPaneCollection, assemblies);
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