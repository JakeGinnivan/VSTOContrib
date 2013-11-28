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
        private readonly RibbonViewModelHelper helperUnderTest;

        public the_ribbon_viewmodel_helper()
        {
            helperUnderTest = new RibbonViewModelHelper();
        }

        [Fact]
        public void throws_when_generic_type_not_enum()
        {
            Assert.Throws<ArgumentException>(
                () => helperUnderTest.GetRibbonTypesFor<TestStruct>(typeof (TestRibbonViewModel)));
        }

        [Fact]
        public void returns_single_ribbon_type_value()
        {
            var results = helperUnderTest.GetRibbonTypesFor<TestRibbonTypes>(typeof(TestRibbonViewModel));

            Assert.Equal(TestRibbonTypes.RibbonType1, results.Single());
        }

        [Fact]
        public void returns_multiple_ribbon_type_value()
        {
            var results = helperUnderTest.GetRibbonTypesFor<TestRibbonTypes>(typeof(TestRibbonViewModel2)).ToList();

            Assert.Equal(TestRibbonTypes.RibbonType2, results[0]);
            Assert.Equal(TestRibbonTypes.RibbonType3, results[1]);
        }

        [Fact]
        public void allows_default_for_ribbon_type_enums()
        {
            var results = helperUnderTest.GetRibbonTypesFor<TestRibbonTypesWithDefault>(typeof(TestRibbonViewModelWithEnumWithDefault)).ToList();

            Assert.Equal(TestRibbonTypesWithDefault.RibbonType, results[0]);
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
