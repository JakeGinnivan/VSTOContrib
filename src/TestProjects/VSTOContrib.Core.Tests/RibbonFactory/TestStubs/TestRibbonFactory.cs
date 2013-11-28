using System.Reflection;
using Microsoft.Office.Tools;
using VSTOContrib.Core.RibbonFactory.Interfaces;

namespace VSTOContrib.Core.Tests.RibbonFactory.TestStubs
{
    internal class TestRibbonFactory : Core.RibbonFactory.RibbonFactory
    {
        private readonly IViewProvider viewProvider;

        public TestRibbonFactory(
            AddInBase addInBase,
            IViewProvider viewProvider,
            IViewContextProvider contextProvider,
            string fallbackRibbonType,
            params Assembly[] assemblies)
            : base(addInBase, assemblies, contextProvider, fallbackRibbonType)
        {
            this.viewProvider = viewProvider;
        }

        public void ClearCurrent()
        {
            Current = null;
        }

        protected override void ShuttingDown()
        {
            
        }

        protected override void InitialiseRibbonFactoryController(IRibbonFactoryController controller, object application)
        {
            controller.Initialise(viewProvider);
        }
    }
}