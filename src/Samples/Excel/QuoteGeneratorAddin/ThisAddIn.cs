using System;
using System.Windows;
using QuoteGeneratorAddin.Core;
using VSTOContrib.Autofac;
using VSTOContrib.Core.RibbonFactory;
using VSTOContrib.Excel.RibbonFactory;
using Office = Microsoft.Office.Core;

namespace QuoteGeneratorAddin
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            RibbonFactory.Current.SetApplication(Application, this);
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            System.Windows.Application.Current.Shutdown();
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            //Required for WPF support
            if (System.Windows.Application.Current == null)
                new Application { ShutdownMode = ShutdownMode.OnExplicitShutdown };

            return new ExcelRibbonFactory(new AutofacViewModelFactory(new AddinModule()), () => CustomTaskPanes, Globals.Factory, typeof(AddinModule).Assembly);
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }
        
        #endregion
    }
}
