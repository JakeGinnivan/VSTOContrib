using VSTOContrib.Core.RibbonFactory.Internal;
using Xunit;

namespace VSTOContrib.Core.Tests.RibbonFactory.Internal
{
    public class OneToManyCustomTaskPaneAdapterTests
    {
        [Fact]
        public void DisposeShouldDiposeAnyInternalTaskPanes()
        {
            var customTaskPane = new CustomTaskPaneDouble(string.Empty);
            var sut = new OneToManyCustomTaskPaneAdapter("Title");
            sut.Add(new OfficeWin32Window(null, null, null), customTaskPane);

            sut.Dispose();

            Assert.Equal(1, customTaskPane.DisposedCalled);
        }
    }
}