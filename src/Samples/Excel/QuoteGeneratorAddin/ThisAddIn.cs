using System;
using System.Windows;
using Microsoft.Office.Tools;
using QuoteGeneratorAddin.Core;
using VSTOContrib.Autofac;
using VSTOContrib.Core.RibbonFactory;
using VSTOContrib.Excel.RibbonFactory;
using Office = Microsoft.Office.Core;

namespace QuoteGeneratorAddin
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            RibbonFactory.Current.SetApplication(Application, this);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            System.Windows.Application.Current.Shutdown();
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            //Required for WPF support
            if (System.Windows.Application.Current == null)
                new Application { ShutdownMode = ShutdownMode.OnExplicitShutdown };

            return new ExcelRibbonFactory(new AutofacViewModelFactory(new AddinModule()), new Lazy<CustomTaskPaneCollection>(() => CustomTaskPanes), Globals.Factory, typeof(AddinModule).Assembly);
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
