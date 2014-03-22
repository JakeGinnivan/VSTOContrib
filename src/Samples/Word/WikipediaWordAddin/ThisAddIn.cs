using System;
using System.Windows;
using Microsoft.Office.Core;
using VSTOContrib.Autofac;
using VSTOContrib.Word.RibbonFactory;
using WikipediaWordAddin.Core;

namespace WikipediaWordAddin
{
    public partial class ThisAddIn
    {
        private void ThisAddInStartup(object sender, EventArgs e)
        {
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

            return new WordRibbonFactory(this, typeof(AddinModule).Assembly)
            {
                ViewModelFactory = new AutofacViewModelFactory(new AddinModule())
            };
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
