using System;
using System.Windows;
using Microsoft.Office.Core;
using Office.Contrib;
using Office.Contrib.RibbonFactory;
using Office.Contrib.RibbonFactory.Interfaces;
using Office.Word.Contrib.RibbonFactory;
using RazorDocs.Core;

namespace RazorDocs
{
    public partial class ThisAddIn
    {
        private AddinBootstrapper _core;

        private static void ThisAddInStartup(object sender, EventArgs e)
        {
            if (System.Windows.Application.Current == null)
                new Application { ShutdownMode = ShutdownMode.OnExplicitShutdown };

            //Check for updates
            new VstoClickOnceUpdater()
                .CheckForUpdateAsync(
                    r =>
                    {
                        if (r.Updated)
                        {
                            MessageBox.Show("RazorDocs add-in updated");
                        }
                    });
        }

        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new WordRibbonFactory(typeof(AddinBootstrapper).Assembly);
        }

        public override void BeginInit()
        {
            _core = new AddinBootstrapper();
            WordRibbonFactory.SetApplication(Application);
            RibbonFactory.Current.InitialiseFactory(
                t => (IRibbonViewModel)_core.Resolve(t), CustomTaskPanes);
            base.BeginInit();
        }

        private void ThisAddInShutdown(object sender, EventArgs e)
        {
            _core.Dispose();
            //System.Windows.Application.Current.Shutdown();
        }

        private void InternalStartup()
        {
            Startup += ThisAddInStartup;
            Shutdown += ThisAddInShutdown;
        }
    }
}
