using System;
using Microsoft.Office.Tools;
using NSubstitute;

namespace VSTOContrib.Core.Tests.RibbonFactory.TestAddin
{
    class TestFactory : Factory
    {
        public Microsoft.Office.Tools.Ribbon.RibbonFactory GetRibbonFactory()
        {
            return Substitute.For<Microsoft.Office.Tools.Ribbon.RibbonFactory>();
        }

        public AddIn CreateAddIn(IServiceProvider serviceProvider, IHostItemProvider hostItemProvider, string primaryCookie,
            string identifier, object containerComponent, IAddInExtension extension)
        {
            return new TestAddin();
        }

        public CustomTaskPaneCollection CreateCustomTaskPaneCollection(IServiceProvider serviceProvider,
            IHostItemProvider hostItemProvider, string primaryCookie, string identifier, object containerComponent)
        {
            return null;
        }

        public SmartTagCollection CreateSmartTagCollection(IServiceProvider serviceProvider, IHostItemProvider hostItemProvider,
            string primaryCookie, string identifier, object containerComponent)
        {
            return null;
        }
    }
}