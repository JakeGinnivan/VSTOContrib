using System.Reflection;
using Microsoft.Office.Tools;
using VSTOContrib.Core.RibbonFactory.Interfaces;

namespace VSTOContrib.Core.Tests.RibbonFactory.TestStubs
{
    internal class TestRibbonFactory : Core.RibbonFactory.RibbonFactory
    {
        private readonly IOfficeApplicationEvents officeApplicationEvents;

        public TestRibbonFactory(
            AddInBase addInBase,
            IOfficeApplicationEvents officeApplicationEvents,
            IViewContextProvider contextProvider,
            string fallbackRibbonType,
            params Assembly[] assemblies)
            : base(addInBase, UseIfEmpty(assemblies, Assembly.GetCallingAssembly()), contextProvider, 
            officeApplicationEvents, fallbackRibbonType)
        {
            this.officeApplicationEvents = officeApplicationEvents;
        }

        protected override void ShuttingDown()
        {

        }

        protected override void InitialiseRibbonFactoryController(IRibbonFactoryController controller, object application)
        {
            officeApplicationEvents.Initialise(application);
        }
    }
}