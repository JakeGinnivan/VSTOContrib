using System;
using System.Reflection;
using Microsoft.Office.Tools;
using Office.Contrib.RibbonFactory;
using Office.Contrib.RibbonFactory.Interfaces;

namespace Office.Contrib.Tests.RibbonFactory.TestStubs
{
    internal class TestRibbonFactory : Contrib.RibbonFactory.RibbonFactory
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