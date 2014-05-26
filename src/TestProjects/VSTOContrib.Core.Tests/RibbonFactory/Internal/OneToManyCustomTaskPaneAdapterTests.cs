using VSTOContrib.Core.RibbonFactory.Internal;
using Xunit;

namespace VSTOContrib.Core.Tests.RibbonFactory.Internal
{
    public class OneToManyCustomTaskPaneAdapterTests
    {
        [Fact]
        public void DisposeShouldDiposeAnyInternalTaskPanes()
        {
            var original = new CustomTaskPaneDouble(string.Empty);
            var viewContext = new object();
            var sut = new OneToManyCustomTaskPaneAdapter(original, viewContext);

            sut.Dispose();

            Assert.Equal(1, original.DisposedCalled);
        }
    }
}