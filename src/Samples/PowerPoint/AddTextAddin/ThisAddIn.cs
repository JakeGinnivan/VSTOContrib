using System;
using System.Windows;
using AddTextAddin.Core;
using Microsoft.Office.Tools;
using VSTOContrib.Core;
using VSTOContrib.Core.RibbonFactory;
using VSTOContrib.PowerPoint.RibbonFactory;
using Office = Microsoft.Office.Core;

namespace AddTextAddin
{
    public partial class ThisAddIn
    {
        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
	{
		//Required for WPF support
        if (System.Windows.Application.Current == null)
            new Application { ShutdownMode = ShutdownMode.OnExplicitShutdown };

		var assemblyContainingViewModels = typeof (PresentationViewModel).Assembly; // This should be the assembly containing all your VSTOContrib viewmodels
		return new PowerPointRibbonFactory(new DefaultViewModelFactory(), new Lazy<CustomTaskPaneCollection>(() => CustomTaskPanes), Globals.Factory, assemblyContainingViewModels);
	}

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            RibbonFactory.Current.SetApplication(Application, this);
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            System.Windows.Application.Current.Shutdown();
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
