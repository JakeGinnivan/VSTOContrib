using System;
using System.Reflection;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools;
using NSubstitute;
using Office.Contrib;
using Office.Contrib.RibbonFactory;
using Office.Contrib.RibbonFactory.Interfaces;
using Office.Word.Contrib.RibbonFactory;
using Xunit;

namespace Office.Word.Contrib.Tests.RibbonFactory
{
    public class WordRibbonFactoryTests
    {
        [Fact]
        public void TestBootstrapping()
        {
            var ribbonFactory = new TestWordRibbonFactory<WordRibbonType>(typeof (WordRibbonFactoryTests).Assembly);

            var ribbonXml = ribbonFactory.GetCustomUI(WordRibbonType.WordDocument.GetEnumDescription());

            //nSubstitute blows up when mocking Application, will investigate later
            WordRibbonFactory.SetApplication(Substitute.For<Application>());
            ribbonFactory.InitialiseFactory(t => (IRibbonViewModel) Activator.CreateInstance(t),
                                            new CustomTaskPaneCollection());

            Assert.NotNull(ribbonXml);
        }
    }

    public class TestWordRibbonFactory<TRibbonType> : Office.Contrib.RibbonFactory.RibbonFactory where TRibbonType : struct
    {
        private readonly IViewProvider<TRibbonType> _viewProvider;
        private readonly IViewContextProvider _viewContextProvider;

        public TestWordRibbonFactory(
            IViewProvider<TRibbonType> viewProvider, 
            IViewContextProvider viewContextProvider,
            params Assembly[] assemblies)
            : base(new RibbonFactoryImpl<TRibbonType>(assemblies))
        {
            _viewProvider = viewProvider;
            _viewContextProvider = viewContextProvider;
        }

        public override IDisposable InitialiseFactory(
            Func<Type, IRibbonViewModel> ribbonFactory, 
            CustomTaskPaneCollection customTaskPaneCollection)
        {
            return InitialiseFactoryInternal(_viewProvider, ribbonFactory, _viewContextProvider, customTaskPaneCollection);
        }
    }
}