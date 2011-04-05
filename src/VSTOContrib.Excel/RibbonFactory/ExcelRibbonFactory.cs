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
        private static Application _ExcelApplication;
        private ExcelViewProvider _ExcelViewProvider;

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelRibbonFactory"/> class.
        /// </summary>
        /// <param name="assemblies">Assemblies to scan for view models</param>
        public ExcelRibbonFactory(params Assembly[] assemblies)
            : base(new RibbonFactoryImpl<ExcelRibbonType>(assemblies))
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelRibbonFactory"/> class.
        /// </summary>
        /// <param name="viewLocationStrategy">The view location strategy, null for default strategy.</param>
        /// <param name="assemblies">Assemblies to scan for view models</param>
        public ExcelRibbonFactory(
            IViewLocationStrategy viewLocationStrategy,
            params Assembly[] assemblies)
            : base(new RibbonFactoryImpl<ExcelRibbonType>(assemblies, viewLocationStrategy))
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
            if (_ExcelApplication == null)
                throw new InvalidOperationException("Set Excel application instance first trough SetApplication()");

            _ExcelViewProvider = new ExcelViewProvider(_ExcelApplication);
            return InitialiseFactoryInternal(
                _ExcelViewProvider,  
                ribbonFactory,
                new ExcelViewContextProvider(),
                customTaskPaneCollection);
        }

        /// <summary>
        /// Ribbon_s the load.
        /// </summary>
        /// <param name="ribbonUi">The ribbon UI.</param>
        public override void Ribbon_Load(Microsoft.Office.Core.IRibbonUI ribbonUi)
        {
            //Excel does not raise a new document event when we are starting up, and initialise is too soon
            _ExcelViewProvider.RegisterOpenDocuments();
            base.Ribbon_Load(ribbonUi);
        }

        /// <summary>
        /// Sets the Outlook application Instance
        /// </summary>
        /// <param name="application"></param>
        public static void SetApplication(Application application)
        {
            _ExcelApplication = application;
        }
    }
}