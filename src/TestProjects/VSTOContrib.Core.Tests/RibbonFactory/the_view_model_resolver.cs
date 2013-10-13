using System;
using System.Collections.Generic;
using Microsoft.Office.Core;
using Microsoft.Office.Tools;
using NSubstitute;
using VSTOContrib.Core.RibbonFactory;
using VSTOContrib.Core.RibbonFactory.Internal;
using VSTOContrib.Core.Tests.RibbonFactory.TestStubs;
using Xunit;

namespace VSTOContrib.Core.Tests.RibbonFactory
{
    public class the_view_model_resolver
    {
        private readonly Func<IEnumerable<Type>, ViewModelResolver<TestRibbonTypes>> resolverFactory;

        public the_view_model_resolver()
        {
            resolverFactory = vms=>new ViewModelResolver<TestRibbonTypes>(
                vms,
                new RibbonViewModelHelper(),
                new CustomTaskPaneRegister(()=>Substitute.For<CustomTaskPaneCollection>()),
                new TestContextProvider(),
                new TestViewModelFactory(), 
                Substitute.For<Factory>());
        }

        [Fact]
        public void cannot_have_two_view_models_for_same_ribbon_type()
        {
            // arrange
            var viewModels = new [] {typeof(TestViewModel), typeof(TestViewModel2)};

            // act/assert
            Assert.Throws<InvalidOperationException>(()=>resolverFactory(viewModels));
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
