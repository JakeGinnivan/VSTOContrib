using Microsoft.Office.Core;
using System;
using Microsoft.Office.Tools;
using VSTOContrib.Core.RibbonFactory;
using VSTOContrib.Core.RibbonFactory.Interfaces;
using VSTOContrib.Word.RibbonFactory;

namespace TwitterResultsWordAddin
{
    public partial class ThisAddIn
    {
        AddinBootstrapper core;

        private void ThisAddInStartup(object sender, EventArgs e)
        {

        }

        private void ThisAddInShutdown(object sender, EventArgs e)
        {
            core.Dispose();
        }

        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            core = new AddinBootstrapper();
            return new WordRibbonFactory(t => (IRibbonViewModel)core.Resolve(t), new Lazy<CustomTaskPaneCollection>(() => CustomTaskPanes), typeof(AddinBootstrapper).Assembly);
        }

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            RibbonFactory.Current.SetApplication(Application, this);

            Startup += ThisAddInStartup;
            Shutdown += ThisAddInShutdown;
        }
    }
}
