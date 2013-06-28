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
        static Application wordApplication;
        WordViewProvider wordViewProvider;

        /// <summary>
        /// Initializes a new instance of the <see cref="WordRibbonFactory"/> class.
        /// </summary>
        /// <param name="ribbonFactory">A delegate taking a type and returning an instance of the requested type</param>
        /// <param name="customTaskPaneCollection">A delayed resolution instance of the custom task pane collection of your addin 'new Lazy(()=>CustomTaskPaneCollection)'</param>
        /// <param name="assemblies">Assemblies to scan for view models</param>
        public WordRibbonFactory(Func<Type, IRibbonViewModel> ribbonFactory, Lazy<CustomTaskPaneCollection> customTaskPaneCollection, params Assembly[] assemblies)
            : base(new RibbonFactoryController<WordRibbonType>(assemblies, new WordViewContextProvider(), ribbonFactory, customTaskPaneCollection))
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="WordRibbonFactory"/> class.
        /// </summary>
        /// <param name="ribbonFactory">A delegate taking a type and returning an instance of the requested type</param>
        /// <param name="customTaskPaneCollection">A delayed resolution instance of the custom task pane collection of your addin 'new Lazy(()=>CustomTaskPaneCollection)'</param>
        /// <param name="viewLocationStrategy">The view location strategy, null for default strategy.</param>
        /// <param name="assemblies">Assemblies to scan for view models</param>
        public WordRibbonFactory(
            Func<Type, IRibbonViewModel> ribbonFactory,
            Lazy<CustomTaskPaneCollection> customTaskPaneCollection,
            IViewLocationStrategy viewLocationStrategy,
            params Assembly[] assemblies)
            : base(new RibbonFactoryController<WordRibbonType>(assemblies, new WordViewContextProvider(), ribbonFactory, customTaskPaneCollection, viewLocationStrategy))
        {
        }

        /// <summary>
        /// Initialises the ribbon factory.
        /// </summary>
        /// <param name="customTaskPaneCollection">The custom task pane collection.</param>
        /// <returns></returns>
        public override IDisposable InitialiseFactory(
            CustomTaskPaneCollection customTaskPaneCollection)
        {
            if (wordApplication == null)
                throw new InvalidOperationException("Set Word application instance first trough SetApplication()");

            wordViewProvider = new WordViewProvider(wordApplication);
            wordViewProvider.RegisterOpenDocuments();
            return InitialiseFactoryInternal(
                wordViewProvider);
        }

        /// <summary>
        /// Ribbon_s the load.
        /// </summary>
        /// <param name="ribbonUi">The ribbon UI.</param>
        public override void Ribbon_Load(Microsoft.Office.Core.IRibbonUI ribbonUi)
        {
            //Word does not raise a new document event when we are starting up, and initialise is too soon
            if (wordViewProvider != null)
                wordViewProvider.RegisterOpenDocuments();
            base.Ribbon_Load(ribbonUi);
        }

        /// <summary>
        /// Sets the Outlook application Instance
        /// </summary>
        /// <param name="application"></param>
        public static void SetApplication(Application application)
        {
            wordApplication = application;
        }
    }
}