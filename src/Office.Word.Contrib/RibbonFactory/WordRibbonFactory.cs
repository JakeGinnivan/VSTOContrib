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

        /// <summary>
        /// Initializes a new instance of the <see cref="WordRibbonFactory"/> class.
        /// </summary>
        /// <param name="viewLocationStrategy">The view location strategy, use null for default.</param>
        public WordRibbonFactory(IViewLocationStrategy viewLocationStrategy = null)
            : base(viewLocationStrategy)
        {
            
        }

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

            return base.InitialiseFactory(ribbonFactory, customTaskPaneCollection, assemblies);
        }

        /// <summary>
        /// Gets the Word View Provider for the ribbon factory.
        /// </summary>
        /// <returns></returns>
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