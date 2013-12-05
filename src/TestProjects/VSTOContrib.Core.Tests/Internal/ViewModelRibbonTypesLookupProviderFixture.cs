using System;
using VSTOContrib.Core.RibbonFactory;
using VSTOContrib.Core.RibbonFactory.Internal;
using Xunit;

namespace VSTOContrib.Core.Tests.Internal
{
    public class ViewModelRibbonTypesLookupProviderFixture
    {
        readonly ViewModelRibbonTypesLookupProvider sut;

        public ViewModelRibbonTypesLookupProviderFixture()
        {
            sut = new ViewModelRibbonTypesLookupProvider();
        }

        [Fact]
        public void DiscoversMultipleAttributesOnType()
        {
            var types = sut.GetRibbonTypesFor(typeof (MultipleRibbonTypes), null);

            Assert.Equal(2, types.Length);
            Assert.Contains("Test", types);
            Assert.Contains("Test2", types);
        }

        [Fact]
        public void FallsBackToFallbackValueWhenNoTypesDefined()
        {
            var types = sut.GetRibbonTypesFor(typeof(ViewModelRibbonTypesLookupProviderFixture), "Fallback");

            Assert.Equal(1, types.Length);
            Assert.Equal("Fallback", types[0]);
        }

        [Fact]
        public void ThrowsWhenNoFallbackOrTypesDefined()
        {
            Assert.Throws<InvalidOperationException>(
                () => sut.GetRibbonTypesFor(typeof (ViewModelRibbonTypesLookupProviderFixture), null));
        }

        [RibbonViewModel("Test")]
        [RibbonViewModel("Test2")]
        public class MultipleRibbonTypes
        {
        }
    }
}