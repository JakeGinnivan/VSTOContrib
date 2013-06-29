using System;
using System.Reflection;
using Microsoft.Office.Tools;
using NSubstitute;
using VSTOContrib.Core.RibbonFactory;
using VSTOContrib.Core.RibbonFactory.Interfaces;

namespace VSTOContrib.Core.Tests.RibbonFactory
{
    public class TestRibbonFactory<TRibbonType> : Core.RibbonFactory.RibbonFactory where TRibbonType : struct
    {
        private readonly IViewProvider<TRibbonType> viewProvider;

        public TestRibbonFactory(
            Func<Type, IRibbonViewModel> ribbonFactory,
            Lazy<CustomTaskPaneCollection> customTaskPaneCollection,
            IViewProvider<TRibbonType> viewProvider, 
            IViewContextProvider viewContextProvider,
            params Assembly[] assemblies)
            : base(new RibbonFactoryController<TRibbonType>(assemblies, viewContextProvider, ribbonFactory, customTaskPaneCollection, Substitute.For<Factory>()))
        {
            this.viewProvider = viewProvider;
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