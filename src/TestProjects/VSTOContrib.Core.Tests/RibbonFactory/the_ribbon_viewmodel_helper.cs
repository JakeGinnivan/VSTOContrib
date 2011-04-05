using System;
using System.Linq;
using VSTOContrib.Core.RibbonFactory.Internal;
using VSTOContrib.Core.Tests.RibbonFactory.TestStubs;
using Xunit;

namespace VSTOContrib.Core.Tests.RibbonFactory
{
    public class the_ribbon_viewmodel_helper
    {
        private readonly RibbonViewModelHelper _helperUnderTest;

        public the_ribbon_viewmodel_helper()
        {
            _helperUnderTest = new RibbonViewModelHelper();
        }

        [Fact]
        public void throws_when_generic_type_not_enum()
        {
            Assert.Throws<ArgumentException>(
                () => _helperUnderTest.GetRibbonTypesFor<TestStruct>(typeof (TestRibbonViewModel)));
        }

        [Fact]
        public void returns_single_ribbon_type_value()
        {
            var results = _helperUnderTest.GetRibbonTypesFor<TestRibbonTypes>(typeof(TestRibbonViewModel));

            Assert.Equal(TestRibbonTypes.RibbonType1, results.Single());
        }

        [Fact]
        public void returns_multiple_ribbon_type_value()
        {
            var results = _helperUnderTest.GetRibbonTypesFor<TestRibbonTypes>(typeof(TestRibbonViewModel2)).ToList();

            Assert.Equal(TestRibbonTypes.RibbonType2, results[0]);
            Assert.Equal(TestRibbonTypes.RibbonType3, results[1]);
        }
    }
}
