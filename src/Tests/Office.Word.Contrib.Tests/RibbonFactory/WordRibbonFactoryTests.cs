using System;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools;
using NSubstitute;
using Office.Contrib;
using Office.Contrib.RibbonFactory.Interfaces;
using Office.Word.Contrib.RibbonFactory;
using Xunit;

namespace Office.Word.Contrib.Tests.RibbonFactory
{
    public class WordRibbonFactoryTests
    {
        [Fact(Skip = "nSubstitute blows up when mocking Application, will investigate later")]
        public void TestBootstrapping()
        {
            var ribbonFactory = new WordRibbonFactory(typeof (WordRibbonFactoryTests).Assembly);

            var ribbonXml = ribbonFactory.GetCustomUI(WordRibbonType.WordDocument.GetEnumDescription());

            //nSubstitute blows up when mocking Application, will investigate later
            WordRibbonFactory.SetApplication(Substitute.For<Application>());
            ribbonFactory.InitialiseFactory(t => (IRibbonViewModel) Activator.CreateInstance(t),
                                            new CustomTaskPaneCollection());

            Assert.NotNull(ribbonXml);
        }
    }
}