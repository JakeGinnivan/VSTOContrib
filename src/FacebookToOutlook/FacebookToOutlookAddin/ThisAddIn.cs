using System;
using FacebookToOutlook;
using Microsoft.Office.Core;
using Outlook.Utility.RibbonFactory;

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
            return RibbonFactory.Instance;
        }

        private void ThisAddInShutdown(object sender, EventArgs e)
        {
            _core.Dispose();
        }

        private void InternalStartup()
        {
            _core = new AddinCore(Application.Session);
            RibbonFactory.Instance.InitialiseFactory(
                t=>(IRibbonViewModel)_core.Resolve(t),
                Application,
                CustomTaskPanes,
                typeof(AddinCore).Assembly);
            Startup += ThisAddInStartup;
            Shutdown += ThisAddInShutdown;
        }
    }
}
