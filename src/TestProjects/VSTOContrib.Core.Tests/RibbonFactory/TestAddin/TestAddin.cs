using System;
using System.ComponentModel;
using System.Windows.Forms;
using Microsoft.Office.Tools;

namespace VSTOContrib.Core.Tests.RibbonFactory.TestAddin
{
    public class TestAddin : AddIn
    {
        public void Dispose() { }

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

        public virtual void OnStartup()
        {
            var handler = Startup;
            if (handler != null) handler(this, EventArgs.Empty);
        }

        public event EventHandler Shutdown;
        public event EventHandler BindingContextChanged;
    }
}