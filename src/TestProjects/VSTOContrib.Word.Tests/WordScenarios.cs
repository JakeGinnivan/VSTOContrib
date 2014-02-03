using Microsoft.Office.Core;
using Office.TestDoubles;
using VSTOContrib.Core;
using VSTOContrib.Core.Tests.RibbonFactory.TestAddin;
using VSTOContrib.Word.RibbonFactory;
using Word.TestDoubles;
using Xunit;

namespace VSTOContrib.Word.Tests
{
    public class WordScenarios
    {
        readonly TestAddInBase testAddInBase;
        readonly WordRibbonFactory sut;
        readonly IRibbonUI ribbonUI;
        readonly ApplicationTestDouble application;
        readonly DocumentTestDouble document;

        public WordScenarios()
        {
            application = new ApplicationTestDouble();
            ribbonUI = new RibbonUITestDouble();
            VstoContribLog.ToTrace();
            VstoContribLog.SetLevel(VstoContribLogLevel.Debug);
            testAddInBase = new TestAddInBase();
            sut = new WordRibbonFactory(testAddInBase);
        }

        [Fact]
        public void OpenWord2013()
        {
            sut.GetCustomUI("Microsoft.Word.Document");
            sut.Ribbon_Load(ribbonUI);

            testAddInBase.SetApplication(application);
            testAddInBase.RaiseStartupEvent();
        } 
    }
}