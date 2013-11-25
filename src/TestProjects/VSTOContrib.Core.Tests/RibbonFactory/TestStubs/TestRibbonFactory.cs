using System;
using System.Reflection;
using Microsoft.Office.Tools;
using NSubstitute;
using VSTOContrib.Core.RibbonFactory;
using VSTOContrib.Core.RibbonFactory.Interfaces;

namespace VSTOContrib.Core.Tests.RibbonFactory.TestStubs
{
    internal class TestRibbonFactory : Core.RibbonFactory.RibbonFactory
    {
        private readonly IViewProvider<TestRibbonTypes> viewProvider;

        public TestRibbonFactory(
            IViewModelFactory viewModelFactory, 
            Func<CustomTaskPaneCollection> customTaskPaneCollection,
            IViewProvider<TestRibbonTypes> viewProvider,
            IViewContextProvider contextProvider,
            params Assembly[] assemblies)
            : base(new RibbonFactoryController<TestRibbonTypes>(assemblies, contextProvider, viewModelFactory, customTaskPaneCollection, Substitute.For<Factory>(), new DefaultViewLocationStrategy()))
        {
            this.viewProvider = viewProvider;
        }

        public void ClearCurrent()
        {
            Current = null;
        }

        protected override void ShuttingDown()
        {
            
        }

        protected override void InitialiseRibbonFactoryController(IRibbonFactoryController controller, object application)
        {
            controller.Initialise(viewProvider);
        }
    }
}