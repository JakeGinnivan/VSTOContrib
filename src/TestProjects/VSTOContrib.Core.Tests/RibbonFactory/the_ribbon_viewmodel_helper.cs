using System;
using System.ComponentModel;
using System.Linq;
using VSTOContrib.Core.RibbonFactory.Internal;
using VSTOContrib.Core.Tests.RibbonFactory.TestStubs;
using Xunit;

namespace VSTOContrib.Core.Tests.RibbonFactory
{
    public class the_ribbon_viewmodel_helper
    {
        private readonly ViewModelRibbonTypesLookupProvider helperUnderTest;

        public the_ribbon_viewmodel_helper()
        {
            helperUnderTest = new ViewModelRibbonTypesLookupProvider();
        }

        [Fact]
        public void returns_single_ribbon_type_value()
        {
            var results = helperUnderTest.GetRibbonTypesFor(typeof(TestRibbonViewModel), null);

            Assert.Equal(TestRibbonTypes.RibbonType1.GetEnumDescription(), results.Single());
        }

        [Fact]
        public void returns_multiple_ribbon_type_value()
        {
            var results = helperUnderTest.GetRibbonTypesFor(typeof(TestRibbonViewModel2), null).ToList();

            Assert.Equal(TestRibbonTypes.RibbonType2.GetEnumDescription(), results[0]);
            Assert.Equal(TestRibbonTypes.RibbonType3.GetEnumDescription(), results[1]);
        }

        [Fact]
        public void allows_default_for_ribbon_type_enums()
        {
            var results = helperUnderTest.GetRibbonTypesFor(typeof(TestRibbonViewModelWithEnumWithDefault), "Foo").ToList();

            Assert.Equal("Foo", results[0]);
        }
    }

    public class TestRibbonViewModelWithEnumWithDefault
    {
    }

    [DefaultValue(RibbonType)]
    public enum TestRibbonTypesWithDefault
    {
        RibbonType
    }
}
