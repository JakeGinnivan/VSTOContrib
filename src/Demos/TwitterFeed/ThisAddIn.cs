using System;
using System.Windows;
using Microsoft.Office.Core;
using Office.Utility;
using Outlook.Utility.RibbonFactory;
using TwitterFeedCore;

namespace TwitterFeed
{
    public partial class ThisAddIn
    {
        private AddinBootstrapper _core;

        private static void ThisAddInStartup(object sender, EventArgs e)
        {
            if (System.Windows.Application.Current == null)
                new Application {ShutdownMode = ShutdownMode.OnExplicitShutdown};

            //Check for updates
            new VstoClickOnceUpdater()
                .CheckForUpdateAsync(
                    r =>
                        {
                            if (r.Updated)
                            {
                                MessageBox.Show("Twitter feed add-in updated");
                            }
                        });
        }

        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return RibbonFactory.Instance;
        }

        private void ThisAddInShutdown(object sender, EventArgs e)
        {
            _core.Dispose();
            System.Windows.Application.Current.Shutdown();
        }

        private void InternalStartup()
        {
            _core = new AddinBootstrapper();
            RibbonFactory.Instance.InitialiseFactory(
                t => (IRibbonViewModel)_core.Resolve(t),
                Application,
                CustomTaskPanes,
                typeof(AddinBootstrapper).Assembly);
            Startup += ThisAddInStartup;
            Shutdown += ThisAddInShutdown;
        }
    }
}
