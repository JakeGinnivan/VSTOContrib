using System;
using System.Reflection;
using Microsoft.Office.Tools;
using Office.Contrib.RibbonFactory;

namespace Office.Contrib.Tests.RibbonFactory.TestStubs
{
    internal class TestRibbonFactory : Contrib.RibbonFactory.RibbonFactory
    {
        private readonly IViewProvider<TestRibbonTypes> _viewProvider;

        public TestRibbonFactory(IViewProvider<TestRibbonTypes> viewProvider) 
            : base(new RibbonFactoryImpl<TestRibbonTypes>(new DefaultViewLocationStrategy()))
        {
            _viewProvider = viewProvider;
        }

        public override IDisposable InitialiseFactory(
            Func<Type, IRibbonViewModel> ribbonFactory, 
            CustomTaskPaneCollection customTaskPaneCollection,
            params Assembly[] assemblies)
        {
            return InitialiseFactoryInternal(
                _viewProvider, ribbonFactory, 
                customTaskPaneCollection, assemblies);
        }

        public void ClearCurrent()
        {
            Current = null;
        }
    }
}