using System;
using System.Windows;
using Microsoft.Office.Core;
using Microsoft.Office.Tools;
using VSTOContrib.Core.RibbonFactory;
using VSTOContrib.Core.RibbonFactory.Interfaces;
using VSTOContrib.Word.RibbonFactory;

namespace WikipediaWordAddin
{
    public partial class ThisAddIn
    {
        AddinBootstrapper core;

        private void ThisAddInStartup(object sender, EventArgs e)
        {
            RibbonFactory.Current.SetApplication(Application, this);
        }

        private void ThisAddInShutdown(object sender, EventArgs e)
        {
            core.Dispose();
            System.Windows.Application.Current.Shutdown();
        }

        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            //Required for WPF support
            if (System.Windows.Application.Current == null)
                new Application { ShutdownMode = ShutdownMode.OnExplicitShutdown };

            core = new AddinBootstrapper();
            return new WordRibbonFactory(t => (IRibbonViewModel)core.Resolve(t), new Lazy<CustomTaskPaneCollection>(() => CustomTaskPanes), Globals.Factory, typeof(AddinBootstrapper).Assembly);
        }

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            Startup += ThisAddInStartup;
            Shutdown += ThisAddInShutdown;
        }
    }
}
