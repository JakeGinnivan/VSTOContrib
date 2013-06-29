using Microsoft.Office.Tools;
using NSubstitute;

namespace VSTOContrib.Core.Tests.RibbonFactory
{
    public class AddInBaseFactory
    {
        private class TestAddIn : AddInBase
        {
            public TestAddIn() : base(Substitute.For<Factory>(), null, null, null)
            {
            }
        }

        public static AddInBase Create()
        {
            return new TestAddIn();
        }
    }
}