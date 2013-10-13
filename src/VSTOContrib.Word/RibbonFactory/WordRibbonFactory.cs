using System;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools;
using VSTOContrib.Core;
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
        WordViewProvider wordViewProvider;

        /// <summary>
        /// Initializes a new instance of the <see cref="WordRibbonFactory"/> class.
        /// </summary>
        /// <param name="viewModelFactory">A view model factory</param>
        /// <param name="customTaskPaneCollection">A delayed resolution instance of the custom task pane collection of your addin 'new Lazy(()=>CustomTaskPaneCollection)'</param>
        /// <param name="vstoFactory">The VSTO factory (Globals.Factory)</param>
        /// <param name="assemblies">Assemblies to scan for view models</param>
        public WordRibbonFactory(
            IViewModelFactory viewModelFactory,
            Func<object> customTaskPaneCollection, 
            Factory vstoFactory, 
            params Assembly[] assemblies)
            : base(new RibbonFactoryController<WordRibbonType>(assemblies, new WordViewContextProvider(), viewModelFactory, customTaskPaneCollection, vstoFactory))
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="WordRibbonFactory"/> class.
        /// </summary>
        /// <param name="viewModelFactory">A view model factory</param>
        /// <param name="customTaskPaneCollection">A delayed resolution instance of the custom task pane collection of your addin 'new Lazy(()=>CustomTaskPaneCollection)'</param>
        /// <param name="viewLocationStrategy">The view location strategy, null for default strategy.</param>
        /// <param name="vstoFactory">The VSTO factory (Globals.Factory)</param>
        /// <param name="assemblies">Assemblies to scan for view models</param>
        public WordRibbonFactory(
            IViewModelFactory viewModelFactory,
            Func<CustomTaskPaneCollection> customTaskPaneCollection,
            IViewLocationStrategy viewLocationStrategy,
            Factory vstoFactory,
            params Assembly[] assemblies)
            : base(new RibbonFactoryController<WordRibbonType>(assemblies, new WordViewContextProvider(), viewModelFactory, customTaskPaneCollection, vstoFactory, viewLocationStrategy))
        {
        }

        /// <summary>
        /// Called when the add-in is shutting down
        /// </summary>
        protected override void ShuttingDown()
        {
            wordViewProvider.Dispose();
        }

        /// <summary>
        /// Initialisation callback for ribbon factory. The implementation must initialise the controller and 
        /// </summary>
        protected override void InitialiseRibbonFactoryController(IRibbonFactoryController controller, object application)
        {
            wordViewProvider = new WordViewProvider((Application)application);
            controller.Initialise(wordViewProvider);
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
    }
}