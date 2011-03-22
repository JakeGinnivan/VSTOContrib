using System;
using System.Collections.Generic;
using Microsoft.Office.Core;
using Office.Contrib.RibbonFactory;
using Office.Contrib.RibbonFactory.Internal;
using Office.Contrib.Tests.RibbonFactory.TestStubs;
using Xunit;

namespace Office.Contrib.Tests.RibbonFactory
{
    public class the_view_model_resolver
    {
        private readonly Func<IEnumerable<Type>, ViewModelResolver<TestRibbonTypes>> _resolverFactory;

        public the_view_model_resolver()
        {
            _resolverFactory = vms=>new ViewModelResolver<TestRibbonTypes>(
                vms,
                new RibbonViewModelHelper(),
                new CustomTaskPaneRegister());
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
