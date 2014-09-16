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
        readonly ExcelOfficeApplicationEvents excelOfficeApplicationEvents;

        public ExcelRibbonFactory(AddInBase addinBase, params Assembly[] assemblies)
            :this(new ExcelOfficeApplicationEvents(), addinBase, UseIfEmpty(assemblies, Assembly.GetCallingAssembly()))
        {
        }

        private ExcelRibbonFactory(ExcelOfficeApplicationEvents officeApplicationEvents, AddInBase addinBase, Assembly[] assemblies)
            : base(addinBase, assemblies, new ExcelViewContextProvider(),
                officeApplicationEvents, ExcelRibbonType.ExcelWorkbook.GetEnumDescription())
        {
            excelOfficeApplicationEvents = officeApplicationEvents;
        }

        protected override void InitialiseRibbonFactoryController(IRibbonFactoryController controller, object application)
        {
            excelOfficeApplicationEvents.Initialise(application);
            excelOfficeApplicationEvents.RegisterOpenDocuments();
        }

        protected override void ShuttingDown()
        {
            excelOfficeApplicationEvents.Dispose();
        }
    }
}