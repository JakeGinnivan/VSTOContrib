using System;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools;
using VSTOContrib.Core.RibbonFactory;
using VSTOContrib.Core.RibbonFactory.Interfaces;

namespace VSTOContrib.Excel.RibbonFactory
{
    /// <summary>
    /// 
    /// </summary>
    [ComVisible(true)]
    public class ExcelRibbonFactory : Core.RibbonFactory.RibbonFactory
    {
        ExcelViewProvider excelViewProvider;

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelRibbonFactory"/> class.
        /// </summary>
        /// <param name="ribbonFactory">A delegate taking a type and returning an instance of the requested type</param>
        /// <param name="customTaskPaneCollection">A delayed resolution instance of the custom task pane collection of your addin 'new Lazy(()=>CustomTaskPaneCollection)'</param>
        /// <param name="assemblies">Assemblies to scan for view models</param>
        public ExcelRibbonFactory(Func<Type, IRibbonViewModel> ribbonFactory, Lazy<CustomTaskPaneCollection> customTaskPaneCollection, params Assembly[] assemblies)
            : base(new RibbonFactoryController<ExcelRibbonType>(assemblies, new ExcelViewContextProvider(), ribbonFactory, customTaskPaneCollection))
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelRibbonFactory"/> class.
        /// </summary>
        /// <param name="ribbonFactory">A delegate taking a type and returning an instance of the requested type</param>
        /// <param name="customTaskPaneCollection">A delayed resolution instance of the custom task pane collection of your addin 'new Lazy(()=>CustomTaskPaneCollection)'</param>
        /// <param name="viewLocationStrategy">The view location strategy, null for default strategy.</param>
        /// <param name="assemblies">Assemblies to scan for view models</param>
        public ExcelRibbonFactory(
            Func<Type, IRibbonViewModel> ribbonFactory,
            Lazy<CustomTaskPaneCollection> customTaskPaneCollection,
            IViewLocationStrategy viewLocationStrategy,
            params Assembly[] assemblies)
            : base(new RibbonFactoryController<ExcelRibbonType>(assemblies, new ExcelViewContextProvider(), ribbonFactory, customTaskPaneCollection, viewLocationStrategy))
        {
        }

        protected override void InitialiseRibbonFactoryController(IRibbonFactoryController controller, object application)
        {
            var app = (Application) application;
            excelViewProvider = new ExcelViewProvider(app);
            controller.Initialise(excelViewProvider);
        }

        protected override void ShuttingDown()
        {
            excelViewProvider.Dispose();
        }
            

        /// <summary>
        /// Ribbon_s the load.
        /// </summary>
        /// <param name="ribbonUi">The ribbon UI.</param>
        public override void Ribbon_Load(Microsoft.Office.Core.IRibbonUI ribbonUi)
        {
            //Excel does not raise a new document event when we are starting up, and initialise is too soon
            excelViewProvider.RegisterOpenDocuments();
            base.Ribbon_Load(ribbonUi);
        }
    }
}