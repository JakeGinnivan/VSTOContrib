using System;
using System.Windows;
using Microsoft.Office.Core;
using Microsoft.Office.Tools;
using TwitterFeedOutlookAddin.Core;
using VSTOContrib.Core.RibbonFactory;
using VSTOContrib.Core.RibbonFactory.Interfaces;
using VSTOContrib.Outlook.RibbonFactory;

namespace TwitterFeedOutlookAddin
{
    public partial class ThisAddIn
    {
        AddinBootstrapper core;

        private void ThisAddInStartup(object sender, EventArgs e)
        {
            if (System.Windows.Application.Current == null)
                new Application { ShutdownMode = ShutdownMode.OnExplicitShutdown };
        }

        void ThisAddInShutdown(object sender, EventArgs e)
        {
            System.Windows.Application.Current.Shutdown();
        }

        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            core = new AddinBootstrapper();
            return new OutlookRibbonFactory(t => (IRibbonViewModel)core.Resolve(t), new Lazy<CustomTaskPaneCollection>(() => CustomTaskPanes), Globals.Factory, typeof(AddinBootstrapper).Assembly);
        }

        private void InternalStartup()
        {
            RibbonFactory.Current.SetApplication(Application, this);

            Startup += ThisAddInStartup;
            Shutdown += ThisAddInShutdown;
        }
    }
}
