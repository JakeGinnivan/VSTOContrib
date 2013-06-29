using System;
using System.Windows;
using Microsoft.Office.Core;
using Microsoft.Office.Tools;
using VSTOContrib.Autofac;
using VSTOContrib.Core.RibbonFactory;
using VSTOContrib.Word.RibbonFactory;
using WikipediaWordAddin.Core;

namespace WikipediaWordAddin
{
    public partial class ThisAddIn
    {
        private void ThisAddInStartup(object sender, EventArgs e)
        {
            RibbonFactory.Current.SetApplication(Application, this);
        }

        private void ThisAddInShutdown(object sender, EventArgs e)
        {
            System.Windows.Application.Current.Shutdown();
        }

        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            //Required for WPF support
            if (System.Windows.Application.Current == null)
                new Application { ShutdownMode = ShutdownMode.OnExplicitShutdown };

            return new WordRibbonFactory(new AutofacViewModelFactory(new AddinModule()), new Lazy<CustomTaskPaneCollection>(() => CustomTaskPanes), Globals.Factory, typeof(AddinModule).Assembly);
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
