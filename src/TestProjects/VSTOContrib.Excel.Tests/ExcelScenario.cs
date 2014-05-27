using Excel.TestDoubles;
using Office.TestDoubles;
using Shouldly;
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
            var worksheetAndWindow = excelFacade.NewWorksheet();

            var customTaskPanes = testAddInBase.GetCustomTaskPaneCollection();

            customTaskPanes.Count.ShouldBe(1);
            worksheetAndWindow.Item1.Close(false, false, false);
            customTaskPanes.Count.ShouldBe(1);
        }
    }
}