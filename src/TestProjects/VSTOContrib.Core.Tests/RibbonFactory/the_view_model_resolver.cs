using System;
using System.Collections.Generic;
using System.Reflection;
using Microsoft.Office.Core;
using Microsoft.Office.Tools;
using VSTOContrib.Core.RibbonFactory;
using VSTOContrib.Core.RibbonFactory.Internal;
using VSTOContrib.Core.Tests.RibbonFactory.TestAddin;
using VSTOContrib.Core.Tests.RibbonFactory.TestStubs;
using Xunit;

namespace VSTOContrib.Core.Tests.RibbonFactory
{
    public class the_view_model_resolver
    {
        private readonly Func<IEnumerable<Type>, ViewModelResolver> resolverFactory;

        public the_view_model_resolver()
        {
            var testAddInBase = AddInBaseFactory.Create();
            resolverFactory = vms=>new ViewModelResolver(
                vms,
                new CustomTaskPaneRegister(testAddInBase),
                new TestContextProvider(),
                new VstoContribContext(new Assembly[0], testAddInBase, "Foo"),
                new TestViewProvider());
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

        public class MyAddin : AddInBase
        {
            public MyAddin()
                : base(null, null, "AddIn", "ThisAddIn")
            {
            }
        }
    }
}
