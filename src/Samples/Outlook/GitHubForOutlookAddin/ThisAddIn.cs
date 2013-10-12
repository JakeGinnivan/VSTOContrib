using System;
using System.Windows;
using GitHubForOutlook.Core;
using Microsoft.Office.Tools;
using VSTOContrib.Autofac;
using VSTOContrib.Core.RibbonFactory;
using VSTOContrib.Outlook.RibbonFactory;
using Office = Microsoft.Office.Core;

namespace GitHubForOutlookAddin
{
    public partial class ThisAddIn
    {
        private void ThisAddInStartup(object sender, EventArgs e)
        {
            RibbonFactory.Current.SetApplication(Application, this);
        }

        void ThisAddInShutdown(object sender, EventArgs e)
        {
            System.Windows.Application.Current.Shutdown();
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            if (System.Windows.Application.Current == null)
                new Application { ShutdownMode = ShutdownMode.OnExplicitShutdown };

            return new OutlookRibbonFactory(
                new AutofacViewModelFactory(new AddinModule()), 
                new Lazy<CustomTaskPaneCollection>(() => CustomTaskPanes), 
                Globals.Factory, typeof(AddinModule).Assembly);
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
