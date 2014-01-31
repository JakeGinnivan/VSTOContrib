using Microsoft.Office.Core;
using VSTOContrib.Core;
using VSTOContrib.Core.Tests.RibbonFactory.TestAddin;
using VSTOContrib.Word.RibbonFactory;
using Xunit;

namespace VSTOContrib.Word.Tests
{
    public class WordScenarios
    {
        readonly TestAddInBase testAddInBase;
        readonly WordRibbonFactory sut;
        readonly IRibbonUI ribbonUI;
        readonly TestApplication application;
        readonly TestDocumentsCollection documents;
        readonly TestWindow view;
        readonly TestDocument testDocument;

        public WordScenarios()
        {
            application = new TestApplication();
            documents = new TestDocumentsCollection();
            view = new TestWindow();
            testDocument = new TestDocument();
            application.Documents = documents;
            ribbonUI = new TestRibbonUI();
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
            testDocument.Windows.Add(view);
            documents.Add(testDocument);

            testAddInBase.SetApplication(application);
            testAddInBase.RaiseStartupEvent();
        } 
    }
}