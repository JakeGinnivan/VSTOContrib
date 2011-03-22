using System;
using FacebookToOutlook;
using Microsoft.Office.Core;
using Office.Contrib.RibbonFactory;
using Office.Contrib.RibbonFactory.Interfaces;
using Office.Outlook.Contrib.RibbonFactory;

namespace FacebookToOutlookAddin
{
    public partial class ThisAddIn
    {
        private AddinCore _core;

        private static void ThisAddInStartup(object sender, EventArgs e)
        {
        }

        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new OutlookRibbonFactory(typeof(AddinCore).Assembly);
        }

        private void ThisAddInShutdown(object sender, EventArgs e)
        {
            _core.Dispose();
        }

        private void InternalStartup()
        {
            _core = new AddinCore(Application.Session);
            OutlookRibbonFactory.SetApplication(Application);
            RibbonFactory.Current.InitialiseFactory(
                t => (IRibbonViewModel)_core.Resolve(t), CustomTaskPanes);
            Startup += ThisAddInStartup;
            Shutdown += ThisAddInShutdown;
        }
    }
}
