using System;
using System.Reflection;
using Microsoft.Office.Tools;
using VSTOContrib.Core.RibbonFactory;
using VSTOContrib.Core.RibbonFactory.Interfaces;

namespace VSTOContrib.Core.Tests.RibbonFactory.TestStubs
{
    internal class TestRibbonFactory : Core.RibbonFactory.RibbonFactory
    {
        private readonly IViewProvider<TestRibbonTypes> _viewProvider;
        private readonly IViewContextProvider _contextProvider;

        public TestRibbonFactory(
            IViewProvider<TestRibbonTypes> viewProvider,
            IViewContextProvider contextProvider,
            params Assembly[] assemblies) 
            : base(new RibbonFactoryImpl<TestRibbonTypes>(assemblies, new DefaultViewLocationStrategy()))
        {
            _viewProvider = viewProvider;
            _contextProvider = contextProvider;
        }

        public override IDisposable InitialiseFactory(
            Func<Type, IRibbonViewModel> ribbonFactory,
            CustomTaskPaneCollection customTaskPaneCollection)
        {
            return InitialiseFactoryInternal(_viewProvider, ribbonFactory, _contextProvider, customTaskPaneCollection);
        }

        public void ClearCurrent()
        {
            Current = null;
        }
    }
}