using System;
using FacebookToOutlook;
using Microsoft.Office.Core;
using Office.Contrib.RibbonFactory;
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
            return new OutlookRibbonFactory();
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
                t=>(IRibbonViewModel)_core.Resolve(t),
                CustomTaskPanes,
                typeof(AddinCore).Assembly);
            Startup += ThisAddInStartup;
            Shutdown += ThisAddInShutdown;
        }
    }
}
