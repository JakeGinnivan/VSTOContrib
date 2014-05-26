using Excel.TestDoubles;
using Office.TestDoubles;
using VSTOContrib.Core;
using VSTOContrib.Core.Tests.RibbonFactory.TestAddin;
using VSTOContrib.Excel.RibbonFactory;
using Xunit;

namespace VSTOContrib.Excel.Tests
{
    public class ExcelScenario
    {
        readonly TestAddInBase testAddInBase;
        readonly Excel2013Facade excelFacade;
        readonly RibbonUITestDouble ribbonUI;
        readonly ExcelRibbonFactory sut;

        public ExcelScenario()
        {
            excelFacade = new Excel2013Facade();
            ribbonUI = new RibbonUITestDouble();
            VstoContribLog.ToTrace();
            VstoContribLog.SetLevel(VstoContribLogLevel.Debug);
            testAddInBase = new TestAddInBase();
            sut = new ExcelRibbonFactory(testAddInBase);
        }

        [Fact]
        public void OpenExcel2013()
        {
            sut.GetCustomUI("Microsoft.Excel.Workbook");
            sut.Ribbon_Load(ribbonUI);

            testAddInBase.SetApplication(excelFacade.Application);

            // Emulate excel querying the status of the ribbon
            var ribbonControl = new RibbonControlDouble("actionButton", null, "Microsoft.Excel.WorkbookactionButton");
            sut.GetEnabled(ribbonControl);

            testAddInBase.RaiseStartupEvent();

//[Debug] ViewProvider.ViewClosed Raised, View: __ComObject (33639718), Context: NullContext (39530145)
//[Debug] Cleaning up viewmodel for context: NullContext (39530145)
//[Info] ViewModel is SpreadSheetViewModel (21522166)
//[Debug] ViewProvider.NewView Raised, Type: Microsoft.Excel.Workbook, View: __ComObject (33639718), Context: __ComObject (63390070)
//[Info] Building ViewModel of type Microsoft.Excel.Workbook for ribbon Microsoft.Excel.Workbook with context __ComObject (63390070)
//[Debug] Setting RibbonUi [__ComObject (22429634)] for ViewModel
//[Debug] Invalidating showMyAddinPaneButton due to property change notification
        }
    }
}