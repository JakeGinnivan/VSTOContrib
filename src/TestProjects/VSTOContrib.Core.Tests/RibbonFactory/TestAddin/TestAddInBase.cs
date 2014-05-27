using Microsoft.Office.Tools;
using VSTOContrib.Core.Annotations;

namespace VSTOContrib.Core.Tests.RibbonFactory.TestAddin
{
    public class TestAddInBase : AddInBase
    {
        readonly TestFactory factory;

        public TestAddInBase()
            : this(new TestFactory())
        {

        }

        public TestAddInBase(Factory factory)
            : base(factory, null, null, null)
        {
            this.factory = (TestFactory) factory;
            Globals.Factory = factory;
            CustomTaskPanes = new CustomTaskPaneCollectionDouble();
        }

        [UsedImplicitly] internal object Application;
        [UsedImplicitly] internal CustomTaskPaneCollection CustomTaskPanes;

        public CustomTaskPaneCollection GetCustomTaskPaneCollection()
        {
            return CustomTaskPanes;
        }

        public TestAddin TestAddin { get { return (TestAddin)Base; } }

        public void SetApplication(object application)
        {
            Application = application;
        }

        public void RaiseStartupEvent()
        {
            factory.UnderlyingAddIn.OnStartup();
        }
    }
}