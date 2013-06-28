using System;
using System.Reflection;
using Microsoft.Office.Tools;
using VSTOContrib.Core.RibbonFactory;
using VSTOContrib.Core.RibbonFactory.Interfaces;

namespace VSTOContrib.Core.Tests.RibbonFactory.TestStubs
{
    internal class TestRibbonFactory : Core.RibbonFactory.RibbonFactory
    {
        private readonly IViewProvider<TestRibbonTypes> viewProvider;

        public TestRibbonFactory(
            Func<Type, IRibbonViewModel> ribbonFactory, 
            Lazy<CustomTaskPaneCollection> customTaskPaneCollection,
            IViewProvider<TestRibbonTypes> viewProvider,
            IViewContextProvider contextProvider,
            params Assembly[] assemblies)
            : base(new RibbonFactoryController<TestRibbonTypes>(assemblies, contextProvider, ribbonFactory, customTaskPaneCollection, new DefaultViewLocationStrategy()))
        {
            this.viewProvider = viewProvider;
        }

        public override IDisposable InitialiseFactory(
            CustomTaskPaneCollection customTaskPaneCollection)
        {
            return InitialiseFactoryInternal(viewProvider);
        }

        public void ClearCurrent()
        {
            Current = null;
        }
    }
}