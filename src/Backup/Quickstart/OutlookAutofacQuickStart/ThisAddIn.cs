using System;
using System.Windows;
using VSTOContrib.Core.RibbonFactory;
using VSTOContrib.Core.RibbonFactory.Interfaces;
using VSTOContrib.Outlook.RibbonFactory;
using Office = Microsoft.Office.Core;

namespace OutlookAutofacQuickStart
{
    public partial class ThisAddIn
    {
        AddinBootstrapper core;

        void ThisAddInStartup(object sender, EventArgs e)
        {
            // Required for WPF Integration in Outlook
            if (System.Windows.Application.Current == null)
                new Application { ShutdownMode = ShutdownMode.OnExplicitShutdown };

            //To enable background checking of updates uncomment this code
            //new VstoClickOnceUpdater()
            //    .CheckForUpdateAsync(
            //        r =>
            //        {
            //            if (r.Updated)
            //            {
            //                MessageBox.Show(@"Add-in updated");
            //            }
            //        });
        }

        void ThisAddInShutdown(object sender, EventArgs e)
        {
            System.Windows.Application.Current.Shutdown();
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new OutlookRibbonFactory(typeof(AddinBootstrapper).Assembly);
        }

        private void InternalStartup()
        {
            core = new AddinBootstrapper();
            OutlookRibbonFactory.SetApplication(Application);
            RibbonFactory.Current.InitialiseFactory(
                t => (IRibbonViewModel)core.Resolve(t),
                CustomTaskPanes);

            Startup += ThisAddInStartup;
            Shutdown += ThisAddInShutdown;
        }
    }
}
