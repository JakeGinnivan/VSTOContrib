using System.Reflection;
using Microsoft.Office.Tools;
using VSTOContrib.Core.RibbonFactory.Interfaces;

namespace VSTOContrib.Core.Tests.RibbonFactory
{
    public class TestRibbonFactory : Core.RibbonFactory.RibbonFactory
    {
        private readonly IViewProvider viewProvider;

        public TestRibbonFactory(
            AddInBase addInBase,
            IViewProvider viewProvider, 
            IViewContextProvider viewContextProvider,
            string fallbackRibbonType,
            params Assembly[] assemblies)
            : base(addInBase, assemblies, viewContextProvider, fallbackRibbonType)
        {
            this.viewProvider = viewProvider;
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