using System;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools;
using VSTOContrib.Core.RibbonFactory;
using VSTOContrib.Core.RibbonFactory.Interfaces;

namespace VSTOContrib.Word.RibbonFactory
{
    /// <summary>
    /// 
    /// </summary>
    [ComVisible(true)]
    public class WordRibbonFactory : Core.RibbonFactory.RibbonFactory
    {
        private static Application _wordApplication;
        private WordViewProvider _wordViewProvider;

        /// <summary>
        /// Initializes a new instance of the <see cref="WordRibbonFactory"/> class.
        /// </summary>
        /// <param name="assemblies">Assemblies to scan for view models</param>
        public WordRibbonFactory(params Assembly[] assemblies)
            : base(new RibbonFactoryController<WordRibbonType>(assemblies))
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="WordRibbonFactory"/> class.
        /// </summary>
        /// <param name="viewLocationStrategy">The view location strategy, null for default strategy.</param>
        /// <param name="assemblies">Assemblies to scan for view models</param>
        public WordRibbonFactory(
            IViewLocationStrategy viewLocationStrategy,
            params Assembly[] assemblies)
            : base(new RibbonFactoryController<WordRibbonType>(assemblies, viewLocationStrategy))
        {
        }

        /// <summary>
        /// Initialises the ribbon factory.
        /// </summary>
        /// <param name="ribbonFactory">The ribbon factory.</param>
        /// <param name="customTaskPaneCollection">The custom task pane collection.</param>
        /// <returns></returns>
        public override IDisposable InitialiseFactory(
            Func<Type, IRibbonViewModel> ribbonFactory,
            CustomTaskPaneCollection customTaskPaneCollection)
        {
            if (_wordApplication == null)
                throw new InvalidOperationException("Set Word application instance first trough SetApplication()");

            _wordViewProvider = new WordViewProvider(_wordApplication);
            return InitialiseFactoryInternal(
                _wordViewProvider,  
                ribbonFactory,
                new WordViewContextProvider(),
                customTaskPaneCollection);
        }

        /// <summary>
        /// Ribbon_s the load.
        /// </summary>
        /// <param name="ribbonUi">The ribbon UI.</param>
        public override void Ribbon_Load(Microsoft.Office.Core.IRibbonUI ribbonUi)
        {
            //Word does not raise a new document event when we are starting up, and initialise is too soon
            _wordViewProvider.RegisterOpenDocuments();
            base.Ribbon_Load(ribbonUi);
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