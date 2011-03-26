using System;
using System.Reflection;
using Microsoft.Office.Tools;
using Office.Contrib.RibbonFactory;
using Office.Contrib.RibbonFactory.Interfaces;

namespace Office.Word.Contrib.Tests.RibbonFactory
{
    public class TestRibbonFactory<TRibbonType> : Office.Contrib.RibbonFactory.RibbonFactory where TRibbonType : struct
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