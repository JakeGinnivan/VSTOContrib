using System;
using System.Reflection;
using Microsoft.Office.Tools;
using VSTOContrib.Core.RibbonFactory;
using VSTOContrib.Core.RibbonFactory.Interfaces;

namespace VSTOContrib.Core.Tests.RibbonFactory
{
    public class TestRibbonFactory<TRibbonType> : Core.RibbonFactory.RibbonFactory where TRibbonType : struct
    {
        private readonly IViewProvider<TRibbonType> _viewProvider;
        private readonly IViewContextProvider _viewContextProvider;

        public TestRibbonFactory(
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