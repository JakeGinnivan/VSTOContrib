using Microsoft.Office.Tools;

namespace VSTOContrib.Core.Tests.RibbonFactory.TestAddin
{
    public class TestAddInBase : AddInBase
    {
        public TestAddInBase()
            : this(new TestFactory())
        {

        }

        public TestAddInBase(Factory factory)
            : base(new TestFactory(), null, null, null)
        {
            Globals.Factory = factory;
        }

        internal object Application;
        public TestAddin TestAddin { get { return (TestAddin)Base; } }
    }
}