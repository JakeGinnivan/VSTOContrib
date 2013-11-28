using System;
using System.Windows;
using GitHubForOutlook.Core;
using VSTOContrib.Autofac;
using VSTOContrib.Outlook.RibbonFactory;
using Office = Microsoft.Office.Core;

namespace GitHubForOutlookAddin
{
    public partial class ThisAddIn
    {
        private void ThisAddInStartup(object sender, EventArgs e)
        {
        }

        void ThisAddInShutdown(object sender, EventArgs e)
        {
            System.Windows.Application.Current.Shutdown();
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            if (System.Windows.Application.Current == null)
                new Application { ShutdownMode = ShutdownMode.OnExplicitShutdown };

            return new OutlookRibbonFactory(this, typeof (AddinModule).Assembly)
            {
                ViewModelFactory = new AutofacViewModelFactory(new AddinModule())
            };
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            Startup += ThisAddInStartup;
            Shutdown += ThisAddInShutdown;
        }
        
        #endregion
    }
}
