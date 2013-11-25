using System;
using System.ComponentModel;
using System.Windows.Forms;
using Microsoft.Office.Tools;
using NSubstitute;

namespace VSTOContrib.Core.Tests.RibbonFactory
{
    public class AddInBaseFactory
    {
        private class TestAddInBase : AddInBase
        {
            public TestAddInBase()
                : base(new TestFactory(), null, null, null)
            {
            }
        }

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

        class TestAddin : AddIn
        {
            public void Dispose(){}

            // ReSharper disable UnusedAutoPropertyAccessor.Local
            public ISite Site { get; set; }
            public event EventHandler Disposed;
            public ControlBindingsCollection DataBindings { get; private set; }
            public BindingContext BindingContext { get; set; }
            public IAddInExtension DefaultExtension { get; private set; }
            public IAddInExtension Extension { get; private set; }
            public ICachedDataProvider DataHost { get; private set; }
            public IServiceProvider HostContext { get; private set; }
            public IHostItemProvider ItemProvider { get; private set; }
            public object Tag { get; set; }
            public event EventHandler Startup;
            public event EventHandler Shutdown;
            public event EventHandler BindingContextChanged;
        }

        public static AddInBase Create()
        {
            return new TestAddInBase();
        }
    }
}