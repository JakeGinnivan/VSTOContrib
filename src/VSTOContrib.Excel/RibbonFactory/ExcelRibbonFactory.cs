using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Office.Tools;
using VSTOContrib.Core;
using VSTOContrib.Core.RibbonFactory.Interfaces;

namespace VSTOContrib.Excel.RibbonFactory
{
    [ComVisible(true)]
    public class ExcelRibbonFactory : Core.RibbonFactory.RibbonFactory
    {
        readonly ExcelViewProvider excelViewProvider;

        public ExcelRibbonFactory(AddInBase addinBase, params Assembly[] assemblies)
            :this(new ExcelViewProvider(), addinBase, UseIfEmpty(assemblies, Assembly.GetCallingAssembly()))
        {
        }

        private ExcelRibbonFactory(ExcelViewProvider viewProvider, AddInBase addinBase, Assembly[] assemblies)
            : base(addinBase, assemblies, new ExcelViewContextProvider(),
                viewProvider, ExcelRibbonType.ExcelWorkbook.GetEnumDescription())
        {
            excelViewProvider = viewProvider;
        }

        protected override void InitialiseRibbonFactoryController(IRibbonFactoryController controller, object application)
        {
            excelViewProvider.Initialise(application);
            excelViewProvider.RegisterOpenDocuments();
        }

        protected override void ShuttingDown()
        {
            excelViewProvider.Dispose();
        }
    }
}