using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools;
using VSTOContrib.Core;
using VSTOContrib.Core.RibbonFactory.Interfaces;

namespace VSTOContrib.Excel.RibbonFactory
{
    [ComVisible(true)]
    public class ExcelRibbonFactory : Core.RibbonFactory.RibbonFactory
    {
        ExcelViewProvider excelViewProvider;

        public ExcelRibbonFactory(AddInBase addinBase, params Assembly[] assemblies)
            : base(addinBase, UseIfEmpty(assemblies, Assembly.GetCallingAssembly()), new ExcelViewContextProvider(), ExcelRibbonType.ExcelWorkbook.GetEnumDescription())
        {
        }

        protected override void InitialiseRibbonFactoryController(IRibbonFactoryController controller, object application)
        {
            var app = (Application) application;
            excelViewProvider = new ExcelViewProvider(app);
            controller.Initialise(excelViewProvider);
            excelViewProvider.RegisterOpenDocuments();
        }

        protected override void ShuttingDown()
        {
            excelViewProvider.Dispose();
        }

        public override void Ribbon_Load(Microsoft.Office.Core.IRibbonUI ribbonUi)
        {
            //Excel does not raise a new document event when we are starting up, and initialise is too soon
            if (excelViewProvider != null)
                excelViewProvider.RegisterOpenDocuments();
            base.Ribbon_Load(ribbonUi);
        }
    }
}