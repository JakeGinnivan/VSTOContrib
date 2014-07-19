using System;
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
        readonly Word2013Facade wordFacade;
        static int counter = 1;

        public WordScenarios()
        {
            OfficeWin32Window.ResolveWindowHandle = o => new IntPtr(counter++);

            wordFacade = new Word2013Facade();
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
            var documentAndWindow = wordFacade.NewDocumentInNewWindow();

            testAddInBase.SetApplication(wordFacade.Application);
            testAddInBase.RaiseStartupEvent();
        } 
    }
}