First install VSTOContrib.Word if you haven't already (you need to install separately for installer to run correctly)

To get started your ThisAddIn.cs should look something like this


using System;
using System.Windows;
using VSTOContrib.Core.RibbonFactory;
using VSTOContrib.Core.RibbonFactory.Interfaces;
using VSTOContrib.Word.RibbonFactory;
using Office = Microsoft.Office.Core;

namespace WordQuickstart
{
    public partial class ThisAddIn
    {
        AddinBootstrapper core;

        private void ThisAddInStartup(object sender, EventArgs e)
        {
            // Required for WPF Integration in Word
            if (System.Windows.Application.Current == null)
                new Application();
            System.Windows.Application.Current.ShutdownMode = ShutdownMode.OnExplicitShutdown;

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

        private void ThisAddInShutdown(object sender, EventArgs e)
        {
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new WordRibbonFactory(typeof(AddinBootstrapper).Assembly);
        }

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            core = new AddinBootstrapper();
            WordRibbonFactory.SetApplication(Application);
            RibbonFactory.Current.InitialiseFactory(
                t => (IRibbonViewModel)core.Resolve(t),
                CustomTaskPanes);

            Startup += ThisAddInStartup;
            Shutdown += ThisAddInShutdown;
        }
    }
}
