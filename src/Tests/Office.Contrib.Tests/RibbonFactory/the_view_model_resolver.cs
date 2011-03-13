using System;
using System.Collections.Generic;
using Microsoft.Office.Core;
using Microsoft.Office.Tools;
using NSubstitute;
using Office.Contrib.RibbonFactory;
using Office.Contrib.Tests.RibbonFactory.TestStubs;
using Xunit;

namespace Office.Contrib.Tests.RibbonFactory
{
    public class the_view_model_resolver
    {
        private readonly IViewProvider<TestRibbonTypes> _viewProvider;
        private readonly Func<IEnumerable<Type>, ViewModelResolver<TestRibbonTypes>> _resolverFactory;

        public the_view_model_resolver()
        {
            _viewProvider = Substitute.For<IViewProvider<TestRibbonTypes>>();
            _resolverFactory = vms=>new ViewModelResolver<TestRibbonTypes>(
                vms,
                t=>(IRibbonViewModel)Activator.CreateInstance(t),
                new RibbonViewModelHelper(),
                new CustomTaskPaneCollection(),
                _viewProvider);
        }

        [Fact]
        public void cannot_have_two_view_models_for_same_ribbon()
        {
            // arrange
            var viewModels = new [] {typeof(TestViewModel), typeof(TestViewModel2)};

            // act/assert
            Assert.Throws<InvalidOperationException>(()=>_resolverFactory(viewModels));
        }

        [RibbonViewModel(TestRibbonTypes.RibbonType1)]
        public class TestViewModel 
        {
            public IRibbonUI RibbonUi { get; set; }
            public void Displayed(object context)
            {
            }

            public void Cleanup()
            {
            }
        }
        [RibbonViewModel(TestRibbonTypes.RibbonType1)]
        public class TestViewModel2 
        {
            public IRibbonUI RibbonUi { get; set; }
            public void Displayed(object context)
            {
            }

            public void Cleanup()
            {
            }
        }
    }
}
