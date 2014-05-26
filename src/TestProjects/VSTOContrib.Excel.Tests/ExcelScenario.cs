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
            sut.GetCustomUI("Microsoft.Word.Document");
            sut.Ribbon_Load(ribbonUI);

            testAddInBase.SetApplication(excelFacade.Application);
            testAddInBase.RaiseStartupEvent();
        }
    }

    public class Excel2013Facade
    {
        public Excel2013Facade()
        {
            Application = new ApplicationTestDouble();
        }

        public ApplicationTestDouble Application { get; private set; }
    }
}